const runtimeContext = {
  protocol: window.location.protocol || "",
  isFileMode: (window.location.protocol || "") === "file:"
};

const defaultApiBaseCandidates = [
  `${window.location.origin}/api`,
  "http://127.0.0.1:4174/api",
  "http://localhost:4174/api",
  "http://10.0.54.249:4174/api",
  "http://172.24.208.1:4174/api",
  "http://172.27.64.1:4174/api"
];

const draftStorage = {
  key: "game-pnl-workbench.v2.9.drafts",
  legacyKey: "game-pnl-workbench.v2.2.drafts",
  version: 5,
  autoSaveDelay: 500
};

const reportStyles = `
  :root {
    color-scheme: light;
    --bg: #f5f1e8;
    --panel: #fffdf8;
    --ink: #1d2a3a;
    --muted: #66758b;
    --line: rgba(29, 42, 58, 0.12);
    --accent: #0f6cbd;
    --accent2: #f77f00;
    --good: #1f8f6a;
    --bad: #c03b33;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    padding: 32px;
    font-family: "Microsoft YaHei UI", "Segoe UI", sans-serif;
    color: var(--ink);
    background: linear-gradient(160deg, #f0e7da 0%, #faf7f1 45%, #efe5d8 100%);
  }
  .report-shell { max-width: 1280px; margin: 0 auto; display: grid; gap: 20px; }
  .hero, .panel, .metric {
    background: rgba(255,255,255,0.88);
    border: 1px solid var(--line);
    border-radius: 24px;
    box-shadow: 0 16px 44px rgba(29, 42, 58, 0.08);
  }
  .hero, .panel { padding: 24px; }
  .eyebrow { margin: 0 0 10px; color: var(--accent); text-transform: uppercase; letter-spacing: 0.16em; font-size: 12px; }
  h1,h2,h3,p { margin: 0; }
  h1 { font-size: 40px; line-height: 1.08; }
  .subtitle { margin-top: 12px; color: var(--muted); line-height: 1.7; max-width: 780px; }
  .meta { margin-top: 18px; display: flex; flex-wrap: wrap; gap: 10px; }
  .chip { padding: 8px 12px; border-radius: 999px; background: rgba(15, 108, 189, 0.08); color: var(--accent); font-size: 13px; }
  .metric-grid, .split-grid, .chart-grid { display: grid; gap: 14px; }
  .metric-grid { grid-template-columns: repeat(4, minmax(0, 1fr)); }
  .metric { padding: 18px; }
  .metric .label { color: var(--muted); font-size: 13px; }
  .metric .value { margin-top: 8px; font-size: 28px; font-weight: 700; }
  .positive { color: var(--good); }
  .negative { color: var(--bad); }
  .split-grid { grid-template-columns: 1.2fr 0.8fr; }
  .chart-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .section-head { display: grid; gap: 8px; margin-bottom: 14px; }
  .section-head p { color: var(--muted); line-height: 1.6; }
  .report-chart {
    width: 100%;
    height: 320px;
    border-radius: 18px;
    background: linear-gradient(180deg, rgba(15, 108, 189, 0.05), rgba(15, 108, 189, 0.1));
    border: 1px solid rgba(29, 42, 58, 0.06);
  }
  table { width: 100%; border-collapse: collapse; font-size: 13px; background: var(--panel); border-radius: 16px; overflow: hidden; }
  th, td { padding: 10px 12px; border-bottom: 1px solid rgba(29, 42, 58, 0.08); text-align: left; white-space: nowrap; }
  th { background: rgba(15, 108, 189, 0.06); color: var(--muted); font-weight: 600; }
  tbody tr:last-child td { border-bottom: 0; }
  .bar-list { display: grid; gap: 12px; }
  .bar-row { display: grid; grid-template-columns: 120px 1fr 120px; gap: 12px; align-items: center; }
  .bar-track { height: 14px; overflow: hidden; border-radius: 999px; background: rgba(29, 42, 58, 0.08); }
  .bar-fill { height: 100%; border-radius: 999px; background: linear-gradient(90deg, var(--accent), var(--accent2)); }
  .footnote { color: var(--muted); font-size: 12px; line-height: 1.7; }
  @media (max-width: 980px) {
    body { padding: 18px; }
    .metric-grid, .split-grid, .chart-grid { grid-template-columns: 1fr; }
  }
`;

const app = document.getElementById("app");

app.innerHTML = `
  <div class="shell">
    <header class="hero">
      <div class="hero-copy">
        <p class="eyebrow">Game P&L Workbench</p>
        <h1>游戏项目成本预估与损益测算台 V2.9</h1>
        <p class="subtitle">聚焦 ②成本费用预估 与 ③项目损益测算；V2.9 在保留 Excel 化录入体验的基础上，把步骤 3 升级为财务分析视角的结果页。金额单位统一为 USD，缩写按美式记法展示，如 K / M / B。</p>
      </div>
      <div class="hero-status">
        <div class="status-pill">
          <span class="status-dot" id="api-status-dot"></span>
          <div><strong>本地 Excel 解析服务</strong><p id="api-status-text">检测中</p></div>
        </div>
        <div class="status-pill">
          <span class="status-dot muted"></span>
          <div><strong>当前产品范围</strong><p>成本预估 + 项目损益</p></div>
        </div>
      </div>
    </header>
    <main id="main-content"></main>
  </div>
`;

document.getElementById("main-content").outerHTML = `
  <main class="tabbed-workspace">
    <section class="card workflow-nav-card">
      <div class="workflow-nav-head">
        <div class="section-title">
          <h2>测算步骤导航</h2>
          <p>把导入校验、规则设置和结果查看拆开，减少单屏信息压力。</p>
        </div>
        <div class="workspace-actions">
          <div class="workspace-draft-row">
            <label class="draft-slot-field">
              <span>草稿槽位</span>
              <select id="draft-slot-select"></select>
            </label>
            <label class="draft-slot-field">
              <span>草稿名称</span>
              <input id="draft-slot-name-input" type="text" maxlength="40" placeholder="例如：Aniimo-保守版" />
            </label>
          </div>
          <div class="workspace-action-group">
            <button class="hc-action-btn" type="button" id="draft-save-btn">保存当前槽位</button>
            <button class="hc-action-btn" type="button" id="draft-save-as-btn">另存新草稿</button>
            <button class="hc-action-btn" type="button" id="draft-restore-btn">恢复所选草稿</button>
            <button class="hc-action-btn subtle-action-btn" type="button" id="draft-clear-btn">删除所选草稿</button>
          </div>
          <div class="workspace-status" id="draft-status-text">草稿保存检测中</div>
        </div>
      </div>
      <div class="workflow-tabs" id="workflow-tabs">
        <button class="workflow-tab is-active" type="button" data-tab-target="inputs">
          <span class="workflow-step">步骤 1</span>
          <strong>数据导入</strong>
          <p>导入口径、查看分摊后透视与历史样本</p>
        </button>
        <button class="workflow-tab" type="button" data-tab-target="rules">
          <span class="workflow-step">步骤 2</span>
          <strong>费用预估</strong>
          <p>聚焦规则筛选、预期值填写与边界调整</p>
        </button>
        <button class="workflow-tab" type="button" data-tab-target="results">
          <span class="workflow-step">步骤 3</span>
          <strong>结果查看</strong>
          <p>集中查看利润、图表与月度汇总结果</p>
        </button>
      </div>
    </section>

    <section class="tab-panel is-active" data-tab-panel="inputs">
      <section class="card control-card">
        <div class="section-title">
          <h2>步骤 1：导入与边界设置</h2>
          <p>先导入历史实际损益，再查看实际成本趋势、分摊拆解和 What-If 快速测算。</p>
        </div>
        <div class="upload-grid">
          <article class="upload-card">
            <h3>分摊后损益</h3>
            <p>用于保留完整的分摊后透视明细，也可在未上传分摊前/分摊值时直接用于历史总览。</p>
            <div class="inline-upload">
              <label class="upload-btn">导入分摊后<input id="pnl-post-file-input" type="file" accept=".xlsx,.xls" /></label>
              <p class="upload-meta" id="pnl-post-file-meta">未导入，当前使用演示分摊后数据。</p>
            </div>
          </article>
          <article class="upload-card">
            <h3>分摊前损益</h3>
            <p>分摊前口径视为项目组直属费用，主要用于来源拆解和延期测算。</p>
            <div class="inline-upload">
              <label class="upload-btn">导入分摊前<input id="pnl-pre-file-input" type="file" accept=".xlsx,.xls" /></label>
              <p class="upload-meta" id="pnl-pre-file-meta">未导入，当前使用演示分摊前数据。</p>
            </div>
          </article>
          <article class="upload-card">
            <h3>分摊值</h3>
            <p>支持多个 sheet 来源汇总读取；历史总览会把分摊值与分摊前合并成完整成本视图。</p>
            <div class="inline-upload">
              <label class="upload-btn">导入分摊值<input id="pnl-alloc-file-input" type="file" accept=".xlsx,.xls" /></label>
              <p class="upload-meta" id="pnl-alloc-file-meta">未导入，当前使用演示分摊值数据。</p>
            </div>
          </article>
          <article class="upload-card">
            <h3>测算边界与后续口径</h3>
            <p>预测边界仍保留给后续步骤 2 - 3 使用；历史总览会优先展示你刚导入的实际损益。</p>
            <label class="field">
              <span>后续步骤测算口径</span>
              <select id="import-mode-select">
                <option value="post">导入分摊后</option>
                <option value="split">导入分摊前 + 分摊值</option>
              </select>
            </label>
            <div class="assumption-grid step1-boundary-grid">
              <label class="field"><span>预测起始月</span><input id="forecast-start" type="month" /></label>
              <label class="field"><span>预测月数</span><input id="forecast-months" type="number" min="3" max="36" step="1" /></label>
            </div>
          </article>
        </div>
        <div class="upload-grid upload-grid-secondary">
          <article class="upload-card">
            <h3>流水 Excel</h3>
            <p>供步骤 2.2 的 A2 模块使用，和步骤1的历史损益总览分开管理。</p>
            <label class="upload-btn">选择文件<input id="flow-file-input" type="file" accept=".xlsx,.xls" /></label>
            <p class="upload-meta" id="flow-file-meta">未导入，当前使用演示流水数据。</p>
          </article>
        </div>
        <div class="helper-note" id="server-hint">${runtimeContext.isFileMode ? "当前通过本地文件直接打开：已自动关闭后端健康检查与轮询，页面可直接用于交互测试。" : "当前版本支持本地 Excel 解析服务自动检测与重试；若未连接，会继续使用演示数据。"}</div>
      </section>

      <section class="card pivot-card output-card">
        <div class="section-title">
          <h2>历史实际损益概览</h2>
          <p>用最新历史月份拆解费用来源，并观察最近 12 个月的 HC 规模与总支出趋势。</p>
        </div>
        <div class="step1-history-toolbar">
          <label class="field">
            <span>项目预算（USD）</span>
            <input id="history-project-budget" type="text" inputmode="numeric" placeholder="请输入项目预算" />
          </label>
          <label class="field">
            <span>近期月均费用口径</span>
            <select id="history-burn-window">
              <option value="1">最近1个月</option>
              <option value="2">最近2个月</option>
              <option value="3" selected>最近3个月</option>
            </select>
          </label>
        </div>
        <div id="history-overview-summary" class="step1-summary-grid"></div>
        <div class="step1-chart-grid">
          <article class="step1-chart-panel step1-chart-panel-breakdown">
            <div class="chart-header chart-header-breakdown">
              <div>
                <h3 id="history-breakdown-title">研发阶段成本结构透视</h3>
                <p>内环：来源大类（营销 / 直属 / 分摊） | 外环：按最新口径展开的具体费用科目</p>
              </div>
              <div class="breakdown-range-stack">
                <div class="breakdown-mode-switch" id="history-breakdown-mode-switch">
                  <button type="button" class="breakdown-mode-btn is-active" data-breakdown-mode="cumulative">区间累计</button>
                  <button type="button" class="breakdown-mode-btn" data-breakdown-mode="average">月均</button>
                  <button type="button" class="breakdown-mode-btn" data-breakdown-mode="single">单月</button>
                </div>
                <div class="breakdown-range-controls">
                  <label class="field">
                    <span>拆解起始月</span>
                    <select id="history-breakdown-start-month"></select>
                  </label>
                  <label class="field">
                    <span>拆解结束月</span>
                    <select id="history-breakdown-end-month"></select>
                  </label>
                </div>
              </div>
            </div>
            <div id="history-breakdown-panel" class="step1-breakdown-panel"></div>
          </article>
        </div>
      </section>

      <section class="card output-card history-panel">
        <div class="step1-collapsible-head">
          <div class="section-title">
            <h2>分摊后损益透视表</h2>
            <p>默认收起，需要时可展开查看明细。分摊后未上传时，会自动尝试用分摊前 + 分摊值推导。</p>
          </div>
          <button type="button" class="hc-action-btn subtle-action-btn" id="pivot-toggle-btn">展开损益表</button>
        </div>
        <div id="pivot-panel" class="step1-collapsible-panel hidden">
          <div class="table-wrap">
            <table class="pivot-meta">
              <tbody>
                <tr><th>项目名称</th><td id="pivot-display-name">-</td></tr>
                <tr><th>当前期间</th><td id="pivot-max-dt">-</td></tr>
              </tbody>
            </table>
          </div>
          <div class="table-wrap wide pivot-scroll">
            <table class="rules-table pivot-table">
              <thead id="pivot-head"></thead>
              <tbody id="pivot-body"></tbody>
            </table>
          </div>
        </div>
      </section>

      <section class="card output-card step1-scenario-card">
        <div class="section-title">
          <h2>动态交互测算（What-If Analysis）</h2>
          <p>基于近期 Burn Rate、HC 成本和利润率假设，快速预估延期追加投入与发行回本压力。</p>
        </div>
        <div class="step1-whatif-grid">
          <div class="step1-whatif-row">
            <article class="step1-whatif-panel">
              <h3>1. 上线推迟追加投入测算</h3>
              <div class="step1-slider-stack">
                <label class="step1-slider-row">
                  <span>上线推迟月数</span>
                  <input id="delay-months-range" type="range" min="0" max="12" step="1" value="6" />
                  <strong id="delay-months-value">6 个月</strong>
                </label>
                <label class="step1-slider-row">
                  <span>推迟期间 HC 扩张率</span>
                  <input id="delay-hc-growth-range" type="range" min="0" max="100" step="5" value="0" />
                  <strong id="delay-hc-growth-value">+0%</strong>
                </label>
              </div>
              <div class="step1-whatif-metrics">
                <article class="step1-whatif-metric">
                  <span>需追加研发投入</span>
                  <strong id="delay-extra-cost">-</strong>
                </article>
                  <article class="step1-whatif-metric">
                    <span>推迟后总累计成本</span>
                    <strong id="delay-total-cost">-</strong>
                  </article>
              </div>
            </article>
            <article class="step1-whatif-panel">
              <h3>2. 发行回本压力测算</h3>
              <div class="step1-slider-stack">
                <label class="step1-slider-row">
                  <span>期望回本周期</span>
                  <input id="payback-months-range" type="range" min="1" max="24" step="1" value="2" />
                  <strong id="payback-months-value">2 个月</strong>
                </label>
                <label class="step1-slider-row">
                  <span>预估发行利润率（扣除渠道/UA）</span>
                  <input id="gross-margin-range" type="range" min="10" max="90" step="5" value="50" />
                  <strong id="gross-margin-value">50%</strong>
                </label>
              </div>
              <div class="step1-whatif-metrics step1-whatif-metrics-payback">
                <article class="step1-whatif-metric step1-whatif-metric-positive">
                  <span>覆盖回本期研发投入所需总流水</span>
                  <strong id="required-total-revenue">-</strong>
                </article>
                <article class="step1-whatif-metric step1-whatif-metric-positive">
                  <span>覆盖累计研发投入所需总流水</span>
                  <strong id="required-cumulative-total-revenue">-</strong>
                </article>
                <article class="step1-whatif-metric step1-whatif-metric-positive">
                  <span>对应目标月均流水</span>
                  <strong id="required-monthly-revenue">-</strong>
                </article>
              </div>
            </article>
          </div>
        </div>
      </section>
    </section>

    <section class="tab-panel" data-tab-panel="rules">
      <section class="rules-stack">
        <section class="card hc-card" id="hc-card"></section>
        <section class="card revenue-card" id="revenue-card"></section>
        <section class="card cost-card" id="cost-card"></section>
        <section class="card marketing-card" id="marketing-card"></section>
        <section class="card rules-card" id="rules-card"></section>
      </section>
    </section>

    <section class="tab-panel" data-tab-panel="results">
        <section class="card summary-card finance-summary-card">
          <div class="results-head">
            <div class="section-title">
            <h2>步骤 3：财务分析视角损益总览</h2>
            <p>基于步骤 2 的全周期测算结果，集中查看 GL 摘要、利润形成、成本结构与风险情景。</p>
            </div>
          <div class="results-actions">
            <button class="hc-action-btn" type="button" data-export-action="summary-csv">导出月度汇总 CSV</button>
            <button class="hc-action-btn" type="button" data-export-action="bridge-csv">导出利润链路 CSV</button>
            <button class="hc-action-btn" type="button" data-export-action="snapshot-json">导出测算包 JSON</button>
            <button class="hc-action-btn" type="button" data-export-action="report-html">导出分享报告 HTML</button>
          </div>
        </div>
        <div class="metric-grid" id="metric-grid"></div>
        <div id="result-budget-strip" class="result-budget-strip"></div>
        <section class="results-layer-block">
          <div class="section-title results-layer-title">
            <h3>第二层：GL后利润形成与流水构成</h3>
            <p>先看利润形成，再看同一 GL 口径下的地区 / 平台流水构成，帮助快速判断利润来源与收入结构。</p>
          </div>
          <div class="results-two-up">
            <section class="chart-card finance-chart-card">
              <div class="chart-header"><h3>GL后N年利润形成瀑布图</h3><p>按 A2 至 A94 的核心链路展示 GL 窗口内的利润形成结构。</p></div>
              <div id="waterfall-chart" class="echart-panel"></div>
            </section>
            <section class="chart-card finance-chart-card">
            <div class="chart-header chart-header-split">
              <div>
                <h3>GL后N年地区 / 平台流水结构</h3>
                <p>使用步骤 1 导入流水的结构占比，并按步骤 2 当前 A2 预测口径重算各月流水分布。</p>
              </div>
              <div class="segmented-toggle" id="revenue-view-toggle">
                <button class="segmented-toggle-btn is-active" type="button" data-revenue-view="region">地区</button>
                <button class="segmented-toggle-btn" type="button" data-revenue-view="platform">平台</button>
              </div>
            </div>
            <div id="revenue-structure-chart" class="echart-panel"></div>
            </section>
          </div>
        </section>
        <section class="results-layer-block">
          <div class="section-title results-layer-title">
            <h3>第三层：成本结构演变</h3>
            <p>成本结构从研发阶段开始观察，因此单独成层展示；横轴按项目阶段逐月标注为研发 nM 或 GL nM，便于识别阶段变化。</p>
          </div>
          <section class="chart-card finance-chart-card finance-chart-card-wide">
            <div class="chart-header"><h3>研发阶段至 GL 后成本结构演变</h3><p>按研发人力、内包与外包、营销与买量、其他经营成本观察从研发阶段至 GL 后的成本变化。</p></div>
            <div id="cost-structure-chart" class="echart-panel echart-panel-tall"></div>
          </section>
        </section>
        <section class="chart-card finance-chart-card stress-chart-card">
          <div class="chart-header">
            <div>
              <h3>风险情景互动测试</h3>
              <p>收入下移会同步下调收入驱动型费用；UA 上浮仅影响 A11 / A11.1 / A11.2，研发与品牌预算保持不变。</p>
            </div>
          </div>
          <div class="stress-control-grid">
            <article class="stress-control-card">
              <div class="stress-control-head">
                <strong>场景 1：收入下移</strong>
                <span>联动收入驱动型费用</span>
              </div>
              <label class="step1-slider-row stress-slider-row">
                <span>收入下移比例</span>
                <input id="stress-revenue-down-range" type="range" min="0" max="60" step="5" value="0" />
                <strong id="stress-revenue-down-value">0%</strong>
              </label>
              <div class="stress-metric-grid">
                <article class="stress-metric">
                  <span>GL 静态回收期</span>
                  <strong id="stress-revenue-payback">-</strong>
                </article>
                <article class="stress-metric">
                  <span>GL 窗口净利润</span>
                  <strong id="stress-revenue-profit">-</strong>
                </article>
                <article class="stress-metric">
                  <span>项目利润率</span>
                  <strong id="stress-revenue-margin">-</strong>
                </article>
              </div>
            </article>
            <article class="stress-control-card">
              <div class="stress-control-head">
                <strong>场景 2：UA 上浮</strong>
                <span>仅调整 A11 / A11.1 / A11.2</span>
              </div>
              <label class="step1-slider-row stress-slider-row">
                <span>UA 上浮比例</span>
                <input id="stress-ua-up-range" type="range" min="0" max="100" step="5" value="0" />
                <strong id="stress-ua-up-value">0%</strong>
              </label>
              <div class="stress-metric-grid">
                <article class="stress-metric">
                  <span>GL 静态回收期</span>
                  <strong id="stress-ua-payback">-</strong>
                </article>
                <article class="stress-metric">
                  <span>GL 窗口净利润</span>
                  <strong id="stress-ua-profit">-</strong>
                </article>
                <article class="stress-metric">
                  <span>项目利润率</span>
                  <strong id="stress-ua-margin">-</strong>
                </article>
              </div>
            </article>
          </div>
          <div id="stress-test-chart" class="echart-panel stress-curve-panel"></div>
        </section>
      </section>

      <section class="output-grid" id="output-grid"></section>
    </section>
  </main>
`;

document.getElementById("hc-card").innerHTML = `
  <div class="module-header">
    <div class="section-title">
      <h2>步骤 2.1：HC 预估</h2>
      <p>V2.0 改为预算模板式布局：左侧维护科目与规则字段，右侧按月份展开，保留贴近 Excel 的录入、复制与填充体验。</p>
    </div>
    <div class="module-pill">V2.0 Excel 形态</div>
  </div>
  <section class="hc-module" id="hc-module"></section>
`;

document.getElementById("revenue-card").innerHTML = `
  <div class="module-header">
    <div class="section-title">
      <h2>步骤 2.2：项目流水和费率</h2>
      <p>V2.1 按 Excel 中步骤 2.2 的 A2/A3/A4/A5 科目重构：左侧识别模板字段，右侧按月预估，并自动联动 A5 净收入。</p>
    </div>
    <div class="module-pill">V2.1 Excel 形态</div>
  </div>
  <section class="revenue-module" id="revenue-module"></section>
`;

document.getElementById("cost-card").innerHTML = `
  <div class="module-header">
    <div class="section-title">
      <h2>步骤 2.3：成本项预估</h2>
      <p>V2.2 按 Excel 中步骤 2.3 的 A97 / A91 / A6 / A7 / A8 / A54 / A9 科目重构：左侧识别模板字段，右侧按月预估，并自动联动 A9 毛利润。</p>
    </div>
    <div class="module-pill">V2.2 Excel 形态</div>
  </div>
  <section class="cost-module" id="cost-module"></section>
`;

document.getElementById("marketing-card").innerHTML = `
  <div class="module-header">
    <div class="section-title">
      <h2>步骤 2.4：市场费预估</h2>
      <p>V2.9 延续 Excel 中步骤 2.4 的 A10 / A11 / A11.1 / A11.2 / A12 / A12.1 / A99 科目结构：左侧识别模板字段，右侧按月预估，并自动联动 A10 与 A99。</p>
    </div>
    <div class="module-pill">V2.9 Excel 形态</div>
  </div>
  <section class="marketing-module" id="marketing-module"></section>
`;

document.getElementById("rules-card").innerHTML = `
  <div class="rules-header">
    <div class="section-title">
      <h2>步骤 2.5 - 2.7：费用预估模块</h2>
      <p id="rules-subtitle">按人力、外包、行政及其他三个模块拆分展示；基于近 1 年实际数据生成建议值，右侧由使用者填写预期值。</p>
    </div>
  <div class="filters rules-toolbar">
      <label class="filter-field rules-search-field"><span>快速搜索</span><input id="rule-search-input" type="search" placeholder="搜索科目编号或科目名称" /></label>
      <label class="filter-field"><span>预估口径</span><select id="scope-filter"><option value="post">分摊后</option><option value="pre">分摊前</option><option value="allocated">分摊值</option></select></label>
      <label class="filter-field"><span>费用来源</span><select id="source-filter"><option value="all">全部</option><option value="项目组直属">项目组直属</option><option value="分摊值">分摊值</option></select></label>
      <label class="filter-field"><span>规则类型</span><select id="rule-filter"><option value="all">全部</option><option value="fixed">固定费用</option><option value="variable">变动费用</option></select></label>
    </div>
  </div>
  <div class="rules-stats" id="rules-stats"></div>
  <div id="rules-body" class="rules-module-stack"></div>
`;

console.log("app scaffold ready");

document.getElementById("output-grid").insertAdjacentHTML(
  "beforeend",
  `
    <section class="card output-card">
      <div class="section-title">
        <h2>自然年财务报表</h2>
        <p>按自然年横向展开，纵向查看核心损益科目，适合年度预算与经营复盘。</p>
      </div>
      <div class="table-wrap wide finance-report-wrap" id="annual-report-wrap"></div>
    </section>
    <section class="card output-card">
      <div class="section-title">
        <h2>GL年财务报表</h2>
        <p>按 GL1Y - GL5Y 横向展开，纵向查看核心损益科目，适合发行后经营跟踪与滚动判断。</p>
      </div>
      <div class="table-wrap wide finance-report-wrap" id="gl-report-wrap"></div>
    </section>
  `
);

const templateAccounts = [
  { code: "A1", name: "项目人数总计" },
  { code: "A1.1", name: "项目正式员工人数" },
  { code: "A1.1.1", name: "项目两年内校招人数" },
  { code: "A1.2", name: "项目实习生人数" },
  { code: "A1.3", name: "项目客服人数" },
  { code: "A1.4.1", name: "项目驻场外包人数" },
  { code: "A1.4.2", name: "项目远程外包人数" },
  { code: "A2", name: "项目流水" },
  { code: "A3", name: "增值税" },
  { code: "A4", name: "平台费" },
  { code: "A5", name: "净收入" },
  { code: "A97", name: "其他收入" },
  { code: "A91", name: "内采收入" },
  { code: "A6", name: "分成成本" },
  { code: "A7", name: "服务器成本" },
  { code: "A7.1", name: "其中：制梦服务器成本" },
  { code: "A8", name: "营业税" },
  { code: "A54", name: "版权金" },
  { code: "A9", name: "毛利润" },
  { code: "A10", name: "市场费用合计" },
  { code: "A11", name: "买量投放费用" },
  { code: "A11.1", name: "买量素材制作费用" },
  { code: "A11.2", name: "买量投放测试费用" },
  { code: "A12", name: "品牌营销费用" },
  { code: "A12.1", name: "用户运营费用" },
  { code: "A99", name: "发行利润" },
  { code: "A13", name: "行政办公费总计" },
  { code: "A14", name: "人力成本" },
  { code: "A15", name: "salary" },
  { code: "A16.1", name: "Performance Bonus" },
  { code: "A16.2", name: "Project Bonus" },
  { code: "A17", name: "Salary-Others" },
  { code: "A18", name: "外包费用" },
  { code: "A19", name: "研发CG" },
  { code: "A20", name: "研发美术" },
  { code: "A21.1", name: "用户运营外包-客服" },
  { code: "A21.2", name: "用户运营外包-VIP服务" },
  { code: "A21.3", name: "用户运营外包-内环境审核" },
  { code: "A21.4", name: "用户运营外包-用户体验" },
  { code: "A21.5", name: "用户运营外包-GS游戏支持" },
  { code: "A22", name: "即时翻译服务费（云上曲率）" },
  { code: "A23", name: "用户运营外包-本地化服务" },
  { code: "A24", name: "研发测试外包" },
  { code: "A25", name: "音频外包及其他外包费用" },
  { code: "A26", name: "房租及物业费" },
  { code: "A27", name: "行政办公费其他小计" },
  { code: "A28", name: "办公室装修费用" },
  { code: "A29", name: "员工福利费" },
  { code: "A30", name: "固定资产及无形资产摊销费" },
  { code: "A31", name: "办公用品及服务" },
  { code: "A32", name: "差旅及交通费" },
  { code: "A33", name: "服务咨询费" },
  { code: "A34.1", name: "招聘猎头费" },
  { code: "A34.2", name: "会议相关费用" },
  { code: "A34.8", name: "招待费" },
  { code: "A34.9", name: "商业保险" },
  { code: "A34.10", name: "特许权使用费" },
  { code: "A34.12", name: "税" },
  { code: "A34.13", name: "培训" },
  { code: "A34.14", name: "通讯费" },
  { code: "A34.15", name: "其他" },
  { code: "A53", name: "内包" },
  { code: "A53.1", name: "其中：发行内包" },
  { code: "A53.2", name: "其中：研发内包" },
  { code: "A35", name: "财务费用及其他杂费" },
  { code: "A36", name: "所得税" },
  { code: "A45", name: "营业外支出/收入" },
  { code: "A92", name: "集团管理服务费" }
];

const hcStepLabel = "2.1";
const hcAccountCodes = ["A1", "A1.1", "A1.1.1", "A1.2", "A1.3", "A1.4.1", "A1.4.2"];
const hcMetaOptionSets = {
  estimateMethodRaw: [
    "按实际需求输入或读取excel",
    "费用率",
    "人均成本",
    "月费用额",
    "=A1.1+A1.1.1+A1.2+A1.3+A1.4.1+A1.4.2",
    "=A2-A3-A4",
    "=A5+A97+A91-A6-A7-A8-A54",
    "=A11+A11.1+A11.2+A12+A12.1",
    "=A9-A10",
    "=A15+A16.1+A16.2+A17",
    "=A18+A19+A20+A21+A22+A23+A24+A25+A26+A27+A28",
    "=A29+A34+A46+A47",
    "=A30+A31+A32+A33",
    "=A35+A36+A37+A38+A39+A40+A41+A42+A43+A44+A45",
    "=A48+A49+A50+A51+A52+A53+A54",
    "=A55+A56+A57+A58+A59+A60+A61+A62+A63",
    "=A27-A28-A64-A65-A66-A67",
    "=A68-A69"
  ],
  costDriver: [
    "按需",
    "A2 Gross Revenue",
    "A11 UA",
    "A1.1 HeadCount-FTE + A1.1.1 HeadCount-FTE-2Y + A1.2 HeadCount-Intern",
    "A15 Salary",
    "时间",
    "A1.1 HeadCount-FTE + A1.1.1 HeadCount-FTE-2Y + A1.2 HeadCount-Intern + A1.3 HeadCount-Customerservice + A1.4.1 Headcount-Outsource-inhouse",
    "A37 Net Income"
  ],
  referencePeriodRaw: ["近1个月", "近3个月", "近6个月", "近12个月"]
};
const hcMetaFieldListIds = {
  estimateMethodRaw: "hc-estimate-method-options",
  costDriver: "hc-cost-driver-options"
};
const hcRowBlueprints = {
  A1: {
    displayName: "HeadCount",
    feeType: "计算项",
    estimateMethod: "=A1.1+A1.1.1+A1.2+A1.3+A1.4.1+A1.4.2",
    costDriver: "",
    suggestedLogic: "",
    referencePeriod: "",
    forecastParam: "",
    forecastLogic: "",
    computed: true
  },
  "A1.1": {
    displayName: "HeadCount-FTE",
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel"
  },
  "A1.1.1": {
    displayName: "HeadCount-FTE-2Y",
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel"
  },
  "A1.2": {
    displayName: "HeadCount-Intern",
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel"
  },
  "A1.3": {
    displayName: "HeadCount-Customerservice",
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel"
  },
  "A1.4.1": {
    displayName: "Headcount-Outsource-inhouse",
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel"
  },
  "A1.4.2": {
    displayName: "Headcount-Outsource-online",
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel"
  }
};
const revenueAccountCodes = ["A2", "A3", "A4", "A5"];
const revenueMetaOptionSets = {
  estimateMethodRaw: [
    "按实际需求输入或读取excel",
    "费用率",
    "=A2-A3-A4"
  ],
  costDriver: [
    "按需",
    "A2 Gross Revenue"
  ],
  referencePeriodRaw: ["近1个月", "近3个月", "近6个月", "近12个月"]
};
const revenueMetaFieldListIds = {
  estimateMethodRaw: "revenue-estimate-method-options",
  costDriver: "revenue-cost-driver-options"
};
const revenueRowBlueprints = {
  A2: {
    displayName: "Gross Revenue",
    feeType: "收入",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  },
  A3: {
    displayName: "VAT",
    feeType: "变动费用-收入相关",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近12个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A4: {
    displayName: "Platform Cost",
    feeType: "变动费用-收入相关",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近12个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A5: {
    displayName: "Net Revenue",
    feeType: "计算项",
    estimateMethod: "=A2-A3-A4",
    costDriver: "",
    suggestedLogic: "",
    referencePeriod: "",
    forecastParam: "",
    forecastLogic: "按预估方法公式自动计算",
    computed: true
  }
};
const costStepLabel = "2.3";
const costAccountCodes = ["A97", "A91", "A6", "A7", "A8", "A54", "A9"];
const costMetaOptionSets = {
  estimateMethodRaw: [
    "按实际需求输入或读取excel",
    "费用率",
    "=A5+A97+A91-A6-A7-A8-A54"
  ],
  costDriver: [
    "按需",
    "A2 Gross Revenue"
  ],
  referencePeriodRaw: ["近1个月", "近3个月", "近6个月", "近12个月"]
};
const costMetaFieldListIds = {
  estimateMethodRaw: "cost-estimate-method-options",
  costDriver: "cost-driver-options"
};
const costRowBlueprints = {
  A97: {
    displayName: "Other Revenue",
    feeType: "收入",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  },
  A91: {
    displayName: "Insourcing Revenue",
    feeType: "收入",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  },
  A6: {
    displayName: "Revenue Share",
    feeType: "变动费用-收入相关",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近12个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A7: {
    displayName: "Server Cost",
    feeType: "变动费用-收入相关",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A8: {
    displayName: "Sales Tax",
    feeType: "变动费用-收入相关",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近12个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A54: {
    displayName: "Royalty",
    feeType: "变动费用",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近12个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A9: {
    displayName: "Gross Profit",
    feeType: "计算项",
    estimateMethod: "=A5+A97+A91-A6-A7-A8-A54",
    costDriver: "",
    suggestedLogic: "",
    referencePeriod: "",
    forecastParam: "",
    forecastLogic: "按预估方法公式自动计算",
    computed: true
  }
};
const marketingStepLabel = "2.4";
const marketingAccountCodes = ["A10", "A11", "A11.1", "A11.2", "A12", "A12.1", "A99"];
const marketingMetaOptionSets = {
  estimateMethodRaw: [
    "按实际需求输入或读取excel",
    "费用率",
    "=A11+A11.1+A11.2+A12+A12.1",
    "=A9-A10"
  ],
  costDriver: [
    "按需",
    "A2 Gross Revenue",
    "A11 UA"
  ],
  referencePeriodRaw: ["近1个月", "近3个月", "近6个月", "近12个月"]
};
const marketingMetaFieldListIds = {
  estimateMethodRaw: "marketing-estimate-method-options",
  costDriver: "marketing-cost-driver-options"
};
const marketingRowBlueprints = {
  A10: {
    displayName: "Sales and Marketing",
    feeType: "计算项",
    estimateMethod: "=A11+A11.1+A11.2+A12+A12.1",
    costDriver: "",
    suggestedLogic: "",
    referencePeriod: "",
    forecastParam: "",
    forecastLogic: "按预估方法公式自动计算",
    computed: true
  },
  A11: {
    displayName: "UA",
    feeType: "变动费用-市场类",
    estimateMethod: "费用率",
    costDriver: "A2 Gross Revenue",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  "A11.1": {
    displayName: "UA-Material",
    feeType: "变动费用-市场类",
    estimateMethod: "费用率",
    costDriver: "A11 UA",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  "A11.2": {
    displayName: "UA-Test",
    feeType: "变动费用-市场类",
    estimateMethod: "费用率",
    costDriver: "A11 UA",
    suggestedLogic: "=参考期间费用总额/参考期间成本动因总额",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "=预估参数*成本动因",
    computed: false
  },
  A12: {
    displayName: "Branding",
    feeType: "变动费用-市场类",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  },
  "A12.1": {
    displayName: "User Operation",
    feeType: "变动费用-市场类",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  },
  A99: {
    displayName: "Publishing Profit",
    feeType: "计算项",
    estimateMethod: "=A9-A10",
    costDriver: "",
    suggestedLogic: "",
    referencePeriod: "",
    forecastParam: "",
    forecastLogic: "按预估方法公式自动计算",
    computed: true
  }
};
const ruleModuleDefinitions = [
  {
    key: "labor-estimate",
    title: "步骤 2.5：人力费用预估",
    description: "覆盖 A14 人力成本及四个新增拆分科目。",
    formula: "A14 = A15 salary + A16.1 Performance Bonus + A16.2 Project Bonus + A17 Salary-Others",
    codes: ["A14", "A15", "A16.1", "A16.2", "A17"]
  },
  {
    key: "outsource-estimate",
    title: "步骤 2.6：外包费用预估",
    description: "覆盖 A18~A25 外包相关科目。",
    formula: "A18 = A19 + A20 + A21.1 + A21.2 + A21.3 + A21.4 + A21.5 + A22 + A23 + A24 + A25",
    codes: ["A18", "A19", "A20", "A21.1", "A21.2", "A21.3", "A21.4", "A21.5", "A22", "A23", "A24", "A25"]
  },
  {
    key: "admin-other-estimate",
    title: "步骤 2.7：行政及其他费用预估",
    description: "收纳其余所有行政、内包、财务、税费及其他科目。",
    formula: "包含除步骤 2.3 ~ 2.6 之外的其他全部费用科目。",
    codes: []
  }
];
const ruleModuleCodeMap = new Map(
  ruleModuleDefinitions.flatMap((module) => module.codes.map((code) => [code, module.key]))
);
const linkedRuleDefinitions = {
  A14: {
    label: "A14 = A15 + A16.1 + A16.2 + A17",
    deps: ["A15", "A16.1", "A16.2", "A17"],
    compute: (ctx) => ctx.sum(["A15", "A16.1", "A16.2", "A17"])
  },
  A18: {
    label: "A18 = A19 + A20 + A21.1 + A21.2 + A21.3 + A21.4 + A21.5 + A22 + A23 + A24 + A25",
    deps: ["A19", "A20", "A21.1", "A21.2", "A21.3", "A21.4", "A21.5", "A22", "A23", "A24", "A25"],
    compute: (ctx) => ctx.sum(["A19", "A20", "A21.1", "A21.2", "A21.3", "A21.4", "A21.5", "A22", "A23", "A24", "A25"])
  },
  A13: {
    label: "A13 = 步骤 2.7 行政及其他费用预估内所有叶子科目汇总",
    deps: [],
    compute: (ctx) => ctx.sum(ctx.adminOtherCodes || [])
  },
  A9: {
    label: "A9 = A5 + A97 + A91 - A6 - A7 - A8 - A54",
    deps: ["A97", "A91", "A6", "A7", "A8", "A54"],
    compute: (ctx) => ctx.netRevenue + ctx.get("A97") + ctx.get("A91") - ctx.get("A6") - ctx.get("A7") - ctx.get("A8") - ctx.get("A54")
  },
  A99: {
    label: "A99 = A9 - A10",
    deps: ["A9", "A10"],
    compute: (ctx) => ctx.get("A9") - ctx.get("A10")
  }
};
const linkedRuleCodes = new Set(Object.keys(linkedRuleDefinitions));
const modelIncomeRuleCodes = new Set(["A97", "A91"]);
const ruleDefaultPresets = {
  A97: { type: "variable", driver: "流水", method: "rate", growthRate: 0 },
  A6: { type: "variable", driver: "流水", method: "rate", growthRate: 0 },
  A7: { type: "variable", driver: "流水", method: "rate", growthRate: 0 },
  "A7.1": { type: "variable", driver: "流水", method: "rate", growthRate: 0 },
  A8: { type: "variable", driver: "流水", method: "rate", growthRate: 0 },
  A54: { type: "variable", driver: "流水", method: "rate", growthRate: 0 }
};

const defaultTemplateLayoutRows = [
  { code: "A1", indent: 0 }, { code: "A1.1", indent: 2 }, { code: "A1.1.1", indent: 2 }, { code: "A1.2", indent: 2 },
  { code: "A1.3", indent: 2 },
  { code: "A1.4.1", indent: 2 }, { code: "A1.4.2", indent: 2 }, { code: "A2", indent: 0 }, { code: "A3", indent: 1 },
  { code: "A4", indent: 1 }, { code: "A5", indent: 0 }, { code: "A97", indent: 1 }, { code: "A91", indent: 1 }, { code: "A6", indent: 1 },
  { code: "A7", indent: 1 }, { code: "A7.1", indent: 1 }, { code: "A8", indent: 1 }, { code: "A54", indent: 1 },
  { code: "A9", indent: 0 }, { code: "A10", indent: 0 }, { code: "A11", indent: 1 }, { code: "A11.1", indent: 1 },
  { code: "A11.2", indent: 1 }, { code: "A12", indent: 1 }, { code: "A12.1", indent: 1 }, { code: "A99", indent: 0 },
  { code: "A14", indent: 0 }, { code: "A15", indent: 1 }, { code: "A16.1", indent: 1 }, { code: "A16.2", indent: 1 },
  { code: "A17", indent: 1 }, { code: "A18", indent: 0 }, { code: "A19", indent: 1 }, { code: "A20", indent: 1 },
  { code: "A21.1", indent: 1 }, { code: "A21.2", indent: 1 }, { code: "A21.3", indent: 1 }, { code: "A21.4", indent: 1 },
  { code: "A21.5", indent: 1 }, { code: "A22", indent: 1 }, { code: "A23", indent: 1 }, { code: "A24", indent: 1 },
  { code: "A25", indent: 1 }, { code: "A13", indent: 0 }, { code: "A26", indent: 1 }, { code: "A27", indent: 1 },
  { code: "A28", indent: 2 }, { code: "A29", indent: 2 }, { code: "A30", indent: 2 }, { code: "A31", indent: 2 },
  { code: "A32", indent: 2 }, { code: "A33", indent: 2 }, { code: "A34.1", indent: 2 }, { code: "A34.2", indent: 2 },
  { code: "A34.8", indent: 2 }, { code: "A34.9", indent: 2 }, { code: "A34.10", indent: 2 }, { code: "A34.12", indent: 2 },
  { code: "A34.13", indent: 2 }, { code: "A34.14", indent: 2 }, { code: "A34.15", indent: 2 }, { code: "A53", indent: 0 },
  { code: "A53.1", indent: 1 }, { code: "A53.2", indent: 1 }, { code: "A35", indent: 0 }, { code: "A36", indent: 0 },
  { code: "A45", indent: 0 }, { code: "A92", indent: 0 }, { code: "A94", indent: 0 }
];

const demoActualRows = [
  { month: "2025-07", code: "A14", account: "人力成本", source: "项目组直属", amount: 820000, hc: 39, revenue: 0 },
  { month: "2025-07", code: "A11", account: "买量投放费用", source: "项目组直属", amount: 560000, hc: 39, revenue: 12500000 },
  { month: "2025-07", code: "A7", account: "服务器成本", source: "项目组直属", amount: 120000, hc: 39, revenue: 12500000 },
  { month: "2025-07", code: "A92", account: "集团管理服务费", source: "分摊值", amount: 180000, hc: 39, revenue: 12500000 },
  { month: "2025-08", code: "A14", account: "人力成本", source: "项目组直属", amount: 835000, hc: 40, revenue: 0 },
  { month: "2025-08", code: "A11", account: "买量投放费用", source: "项目组直属", amount: 590000, hc: 40, revenue: 13200000 },
  { month: "2025-08", code: "A7", account: "服务器成本", source: "项目组直属", amount: 126000, hc: 40, revenue: 13200000 },
  { month: "2025-08", code: "A92", account: "集团管理服务费", source: "分摊值", amount: 188000, hc: 40, revenue: 13200000 },
  { month: "2025-09", code: "A14", account: "人力成本", source: "项目组直属", amount: 848000, hc: 41, revenue: 0 },
  { month: "2025-09", code: "A11", account: "买量投放费用", source: "项目组直属", amount: 610000, hc: 41, revenue: 13600000 },
  { month: "2025-09", code: "A7", account: "服务器成本", source: "项目组直属", amount: 131000, hc: 41, revenue: 13600000 },
  { month: "2025-09", code: "A92", account: "集团管理服务费", source: "分摊值", amount: 194000, hc: 41, revenue: 13600000 }
];

const demoPostRows = demoActualRows.map((row) => ({ ...row, source: "分摊后" }));
const demoPreRows = demoActualRows.filter((row) => row.source === "项目组直属").map((row) => ({ ...row, source: "项目组直属" }));
const demoAllocatedRows = demoActualRows.filter((row) => row.source === "分摊值").map((row) => ({ ...row, source: "分摊值" }));

const demoFlowRows = [
  ["2026-01", 13800000], ["2026-02", 14100000], ["2026-03", 14500000], ["2026-04", 14900000],
  ["2026-05", 15100000], ["2026-06", 15500000], ["2026-07", 15900000], ["2026-08", 16200000],
  ["2026-09", 16500000], ["2026-10", 16700000], ["2026-11", 16900000], ["2026-12", 17200000]
].flatMap(([month, revenue], index) => {
  const offshoreWeight = 0.34 + Math.min(index, 6) * 0.01;
  const cnWeight = 1 - offshoreWeight;
  const mobileWeight = 0.68 - Math.min(index, 5) * 0.01;
  const pcWeight = 0.22 + (index % 3) * 0.01;
  const consoleWeight = 1 - mobileWeight - pcWeight;
  const regionPlatformRows = [
    { region: "CN", platform: "移动", ratio: cnWeight * mobileWeight },
    { region: "CN", platform: "PC", ratio: cnWeight * pcWeight },
    { region: "CN", platform: "Console", ratio: cnWeight * consoleWeight },
    { region: "海外", platform: "移动", ratio: offshoreWeight * mobileWeight },
    { region: "海外", platform: "PC", ratio: offshoreWeight * pcWeight },
    { region: "海外", platform: "Console", ratio: offshoreWeight * consoleWeight }
  ];
  return regionPlatformRows.map((item, rowIndex) => ({
    month,
    region: item.region,
    platform: item.platform,
    revenue: rowIndex === regionPlatformRows.length - 1
      ? revenue - sum(regionPlatformRows.slice(0, -1).map((row) => Math.round(revenue * row.ratio)))
      : Math.round(revenue * item.ratio)
  }));
});

let loadedTemplateLayoutNameLookup = new Map();

const state = {
  apiBase: defaultApiBaseCandidates[0],
  apiBaseCandidates: [...defaultApiBaseCandidates],
  apiConnected: false,
  apiStatusText: runtimeContext.isFileMode ? "文件模式" : "检测中",
  apiStatusHint: runtimeContext.isFileMode
    ? "当前从本地文件打开，已关闭后端探测；页面会使用演示数据与手工录入能力。"
    : "正在检查本地 Excel 解析服务…",
  importMode: "post",
  ruleScope: "post",
  actualRows: {
    post: [...demoPostRows],
    pre: [...demoPreRows],
    allocated: [...demoAllocatedRows]
  },
  actualImportFlags: {
    post: false,
    pre: false,
    allocated: false
  },
  flowRows: [...demoFlowRows],
  flowImportName: "",
  flowImportRowCount: demoFlowRows.length,
  flowImportMatchedChannels: [],
  flowImportUnmatchedChannels: [],
  flowImportSourceFormat: "",
  historicalBreakdownMonth: "2025-09",
  historicalBreakdownStartMonth: "2025-09",
  historicalBreakdownEndMonth: "2025-09",
  historicalBreakdownMode: "cumulative",
  historicalBurnWindow: 3,
  projectBudgetAmount: 0,
  pivotExpanded: false,
  budgetAlertLaunchMonth: "",
  budgetAlertMarketingBudget: 0,
  budgetAlertHcGrowthRate: 0,
  budgetAlertExpanded: false,
  delayMonths: 6,
  delayHcGrowthRate: 0,
  paybackMonths: 2,
  grossMarginRate: 0.5,
  resultRevenueView: "region",
  stressRevenueDownRate: 0,
  stressUaUpRate: 0,
  goLiveDate: "",
  revenueAggregateOverrides: {
    rawRevenueValues: {},
    rawVatRateValues: {},
    rawPlatformRateValues: {},
    revenueValues: {},
    vatRateValues: {},
    platformRateValues: {}
  },
  forecastStart: "2025-10",
  forecastMonths: 12,
  hcStart: 1,
  hcForecastRows: [],
  hcSelection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, mode: "cell" },
  hcEditCell: null,
  hcFormulaInsertLock: false,
  hcFillDrag: null,
  revenueSelection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, mode: "cell" },
  revenueEditCell: null,
  revenueFormulaInsertLock: false,
  revenueFillDrag: null,
  revenueForecastRows: [],
  revenueForecastParts: [],
  costSelection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, mode: "cell" },
  costEditCell: null,
  costFormulaInsertLock: false,
  costFillDrag: null,
  costForecastRows: [],
  marketingSelection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, mode: "cell" },
  marketingEditCell: null,
  marketingFormulaInsertLock: false,
  marketingFillDrag: null,
  marketingForecastRows: [],
  rulesSelection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0, mode: "cell" },
  rulesEditCell: null,
  rulesFillDrag: null,
  rules: { post: [], pre: [], allocated: [] },
  templateLayoutRows: normalizeTemplateLayoutRows(defaultTemplateLayoutRows),
  postPivot: { displayName: "演示项目", maxDt: "-", dtColumns: [], rows: [] },
  filters: { source: "all", type: "all", search: "" },
  activeTab: "inputs",
  persistence: {
    available: false,
    hasDraft: false,
    lastSavedAt: "",
    statusText: "草稿保存检测中",
    lastSavedSignature: "",
    dirty: false,
    drafts: [],
    activeDraftId: "",
    activeDraftName: ""
  }
};

const els = {
  apiDot: document.getElementById("api-status-dot"),
  apiText: document.getElementById("api-status-text"),
  serverHint: document.getElementById("server-hint"),
  forecastStart: document.getElementById("forecast-start"),
  forecastMonths: document.getElementById("forecast-months"),
  importMode: document.getElementById("import-mode-select"),
  ruleSearch: document.getElementById("rule-search-input"),
  scopeFilter: document.getElementById("scope-filter"),
  rulesSubtitle: document.getElementById("rules-subtitle"),
  hcModule: document.getElementById("hc-module"),
  revenueModule: document.getElementById("revenue-module"),
  costModule: document.getElementById("cost-module"),
  marketingModule: document.getElementById("marketing-module"),
  rulesStats: document.getElementById("rules-stats"),
  postUploadWrap: document.getElementById("post-upload-wrap"),
  splitUploadWrap: document.getElementById("split-upload-wrap"),
  metricGrid: document.getElementById("metric-grid"),
  resultBudgetStrip: document.getElementById("result-budget-strip"),
  rulesBody: document.getElementById("rules-body"),
  summaryBody: document.getElementById("summary-body"),
  annualReportWrap: document.getElementById("annual-report-wrap"),
  glReportWrap: document.getElementById("gl-report-wrap"),
  profitBridgeBody: document.getElementById("profit-bridge-body"),
  profitBridgeChart: document.getElementById("profit-bridge-chart"),
  historyBody: document.getElementById("history-body"),
  historyCards: document.getElementById("history-cards"),
  waterfallChart: document.getElementById("waterfall-chart"),
  revenueStructureChart: document.getElementById("revenue-structure-chart"),
  costStructureChart: document.getElementById("cost-structure-chart"),
  stressTestChart: document.getElementById("stress-test-chart"),
  revenueViewToggle: document.getElementById("revenue-view-toggle"),
  stressRevenueDownRange: document.getElementById("stress-revenue-down-range"),
  stressRevenueDownValue: document.getElementById("stress-revenue-down-value"),
  stressRevenuePayback: document.getElementById("stress-revenue-payback"),
  stressRevenueProfit: document.getElementById("stress-revenue-profit"),
  stressRevenueMargin: document.getElementById("stress-revenue-margin"),
  stressUaUpRange: document.getElementById("stress-ua-up-range"),
  stressUaUpValue: document.getElementById("stress-ua-up-value"),
  stressUaPayback: document.getElementById("stress-ua-payback"),
  stressUaProfit: document.getElementById("stress-ua-profit"),
  stressUaMargin: document.getElementById("stress-ua-margin"),
  historyBurnWindow: document.getElementById("history-burn-window"),
  historyBreakdownStartMonth: document.getElementById("history-breakdown-start-month"),
  historyBreakdownEndMonth: document.getElementById("history-breakdown-end-month"),
  historyBreakdownModeSwitch: document.getElementById("history-breakdown-mode-switch"),
  historyProjectBudget: document.getElementById("history-project-budget"),
  historyOverviewSummary: document.getElementById("history-overview-summary"),
  historyBreakdownTitle: document.getElementById("history-breakdown-title"),
  historyBreakdownPanel: document.getElementById("history-breakdown-panel"),
  historyTrendChart: document.getElementById("history-trend-chart"),
  pivotToggleButton: document.getElementById("pivot-toggle-btn"),
  pivotPanel: document.getElementById("pivot-panel"),
  pivotDisplayName: document.getElementById("pivot-display-name"),
  pivotMaxDt: document.getElementById("pivot-max-dt"),
  pivotHead: document.getElementById("pivot-head"),
  pivotBody: document.getElementById("pivot-body"),
  budgetAlertLaunchMonth: document.getElementById("budget-alert-launch-month"),
  budgetAlertMarketingBudget: document.getElementById("budget-alert-marketing-budget"),
  budgetAlertHcGrowthRange: document.getElementById("budget-alert-hc-growth-range"),
  budgetAlertHcGrowthValue: document.getElementById("budget-alert-hc-growth-value"),
  budgetAlertStatus: document.getElementById("budget-alert-status"),
  budgetAlertMonthsToLaunch: document.getElementById("budget-alert-months-to-launch"),
  budgetAlertMonthlyCost: document.getElementById("budget-alert-monthly-cost"),
  budgetAlertTotalCost: document.getElementById("budget-alert-total-cost"),
  budgetAlertBudgetGap: document.getElementById("budget-alert-budget-gap"),
  delayMonthsRange: document.getElementById("delay-months-range"),
  delayMonthsValue: document.getElementById("delay-months-value"),
  delayHcGrowthRange: document.getElementById("delay-hc-growth-range"),
  delayHcGrowthValue: document.getElementById("delay-hc-growth-value"),
  delayExtraCost: document.getElementById("delay-extra-cost"),
  delayTotalCost: document.getElementById("delay-total-cost"),
  paybackMonthsRange: document.getElementById("payback-months-range"),
  paybackMonthsValue: document.getElementById("payback-months-value"),
  grossMarginRange: document.getElementById("gross-margin-range"),
  grossMarginValue: document.getElementById("gross-margin-value"),
  requiredTotalRevenue: document.getElementById("required-total-revenue"),
  requiredCumulativeTotalRevenue: document.getElementById("required-cumulative-total-revenue"),
  requiredMonthlyRevenue: document.getElementById("required-monthly-revenue"),
  draftSlotSelect: document.getElementById("draft-slot-select"),
  draftSlotNameInput: document.getElementById("draft-slot-name-input"),
  draftSaveButton: document.getElementById("draft-save-btn"),
  draftSaveAsButton: document.getElementById("draft-save-as-btn"),
  draftRestoreButton: document.getElementById("draft-restore-btn"),
  draftClearButton: document.getElementById("draft-clear-btn"),
  draftStatusText: document.getElementById("draft-status-text"),
  resultActions: [...document.querySelectorAll("[data-export-action]")],
  pnlPostMeta: document.getElementById("pnl-post-file-meta"),
  pnlPreMeta: document.getElementById("pnl-pre-file-meta"),
  pnlAllocMeta: document.getElementById("pnl-alloc-file-meta"),
  flowMeta: document.getElementById("flow-file-meta"),
  tabButtons: [...document.querySelectorAll("[data-tab-target]")],
  tabPanels: [...document.querySelectorAll("[data-tab-panel]")]
};

let historyBreakdownChartInstance = null;
let waterfallChartInstance = null;
let revenueStructureChartInstance = null;
let costStructureChartInstance = null;
let stressTestChartInstance = null;

window.addEventListener("resize", () => {
  historyBreakdownChartInstance?.resize();
  waterfallChartInstance?.resize();
  revenueStructureChartInstance?.resize();
  costStructureChartInstance?.resize();
  stressTestChartInstance?.resize();
});

function addMonths(monthKey, offset) {
  const [year, month] = monthKey.split("-").map(Number);
  const date = new Date(year, month - 1 + offset, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthRange(start, count) {
  return Array.from({ length: count }, (_, index) => addMonths(start, index));
}

function money(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(Math.round(Number(value) || 0));
}

function compactMoney(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    notation: "compact",
    compactDisplay: "short",
    minimumFractionDigits: 0,
    maximumFractionDigits: 1
  }).format(Number(value || 0));
}

function compactDecimal(value, digits = 2) {
  const numeric = Number(value || 0);
  return numeric.toLocaleString("zh-CN", { minimumFractionDigits: 0, maximumFractionDigits: digits });
}

function formatWan(value, digits = 2) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    notation: "compact",
    compactDisplay: "short",
    minimumFractionDigits: 0,
    maximumFractionDigits: digits
  }).format(Number(value || 0));
}

function formatWanPerMonth(value, digits = 2) {
  return `${formatWan(value, digits)}/mo`;
}

function formatWanPerPersonMonth(value, digits = 2) {
  return `${formatWan(value, digits)}/person/mo`;
}

function formatSignedWan(value, digits = 2) {
  const numeric = Number(value) || 0;
  const formatted = formatWan(Math.abs(numeric), digits);
  if (numeric > 0) return `+${formatted}`;
  if (numeric < 0) return `-${formatted}`;
  return formatted;
}

function formatPeopleCount(value) {
  return `${integerWithThousands(value)} 人`;
}

function formatPeopleAverage(value, digits = 1) {
  return `${compactDecimal(value, digits)} 人`;
}

function metricParts(value, unit, digits = 1, divisor = 1) {
  return {
    value: compactDecimal((Number(value) || 0) / divisor, digits),
    unit
  };
}

function formatSignedPercent(value, digits = 1) {
  const numeric = Number(value || 0) * 100;
  return `${numeric >= 0 ? "+" : ""}${numeric.toFixed(digits)}%`;
}

function buildSparklineGeometry(values, width = 104, height = 34, padding = 4) {
  const points = values.map((value) => Number(value || 0)).filter((value) => Number.isFinite(value));
  if (!points.length) {
    const baselineY = height - padding;
    return {
      polyline: `${padding},${baselineY} ${width - padding},${baselineY}`,
      nodes: []
    };
  }
  if (points.length === 1) {
    const y = height / 2;
    return {
      polyline: `${padding},${y} ${width - padding},${y}`,
      nodes: [{ x: width / 2, y, isLatest: true }]
    };
  }
  const minValue = Math.min(...points);
  const maxValue = Math.max(...points);
  const range = maxValue - minValue || 1;
  const nodes = points.map((value, index) => {
    const x = padding + ((width - padding * 2) / Math.max(points.length - 1, 1)) * index;
    const y = height - padding - ((value - minValue) / range) * (height - padding * 2);
    return {
      x: Number(x.toFixed(2)),
      y: Number(y.toFixed(2)),
      isLatest: index === points.length - 1
    };
  });
  return {
    polyline: nodes.map((point) => `${point.x},${point.y}`).join(" "),
    nodes
  };
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function integerWithThousands(value) {
  const rounded = Math.round(Number(value) || 0);
  return new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 0 }).format(rounded);
}

function parseFormattedNumber(value) {
  const sanitized = String(value || "").replace(/[^\d.-]/g, "");
  const numeric = Number(sanitized);
  return Number.isFinite(numeric) ? numeric : 0;
}

function ratio(value) {
  return `${((value || 0) * 100).toFixed(1)}%`;
}

function avg(list) {
  const values = list.filter((item) => Number.isFinite(item));
  return values.length ? values.reduce((sum, item) => sum + item, 0) / values.length : 0;
}

function sum(list) {
  return list.reduce((total, item) => total + item, 0);
}

function normalizeMonth(raw) {
  if (!raw) return "";
  const text = String(raw).trim();
  const fromDate = new Date(text);
  if (!Number.isNaN(fromDate.getTime())) {
    return `${fromDate.getFullYear()}-${String(fromDate.getMonth() + 1).padStart(2, "0")}`;
  }
  const ymd = text.match(/^(\d{4})[-/](\d{1,2})(?:[-/]\d{1,2})?$/);
  if (ymd) return `${ymd[1]}-${ymd[2].padStart(2, "0")}`;
  const ym = text.match(/^(\d{4})(\d{2})$/);
  if (ym) return `${ym[1]}-${ym[2]}`;
  return text;
}

function monthKeyToDate(monthKey) {
  const normalized = normalizeMonth(monthKey);
  if (!normalized) return null;
  const [year, month] = normalized.split("-").map(Number);
  if (!year || !month) return null;
  return new Date(year, month - 1, 1);
}

function parseInputDate(raw) {
  const text = String(raw || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return null;
  const [year, month, day] = text.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function monthDiff(fromDate, toDate) {
  if (!(fromDate instanceof Date) || Number.isNaN(fromDate.getTime()) || !(toDate instanceof Date) || Number.isNaN(toDate.getTime())) return NaN;
  return (toDate.getFullYear() - fromDate.getFullYear()) * 12 + (toDate.getMonth() - fromDate.getMonth());
}

function parseNumericInput(raw) {
  const text = String(raw ?? "").trim().replace(/,/g, "");
  if (!text) return 0;
  const numeric = Number(text.replace(/%$/, ""));
  return Number.isFinite(numeric) ? numeric : 0;
}

function parseRateInput(raw) {
  const text = String(raw ?? "").trim().replace(/,/g, "");
  if (!text) return 0;
  const hasPercent = text.endsWith("%");
  const numeric = Number(text.replace(/%$/, ""));
  if (!Number.isFinite(numeric)) return 0;
  return hasPercent || Math.abs(numeric) > 1 ? numeric / 100 : numeric;
}

function getApiBaseCandidates() {
  return [...new Set([state.apiBase, ...state.apiBaseCandidates].filter(Boolean))];
}

async function fetchFromApi(path, options = {}) {
  let lastError = null;
  for (const base of getApiBaseCandidates()) {
    try {
      const response = await fetch(`${base}${path}`, options);
      state.apiBase = base;
      return { response, base };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("api_unreachable");
}

let autoSaveTimer = null;
let suspendAutoSave = false;

function safeLocalStorage() {
  try {
    const testKey = `${draftStorage.key}.healthcheck`;
    window.localStorage.setItem(testKey, "1");
    window.localStorage.removeItem(testKey);
    return window.localStorage;
  } catch {
    return null;
  }
}

function formatTimestamp(raw) {
  if (!raw) return "";
  const date = new Date(raw);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString("zh-CN", {
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function cloneData(data) {
  return JSON.parse(JSON.stringify(data));
}

function normalizeActualRowsSnapshot(rows) {
  return Array.isArray(rows) ? rows.map((row) => ({
    month: normalizeMonth(row?.month || row?.dt),
    code: String(row?.code || row?.fcst_acc_code || "").trim(),
    account: String(row?.account || row?.fcst_acc_name || "").trim(),
    updateDt: String(row?.updateDt || row?.update_dt || "").trim(),
    source: String(row?.source || "").trim(),
    amount: Number(row?.amount || 0),
    hc: Number(row?.hc || 0),
    revenue: Number(row?.revenue || 0)
  })).filter((row) => row.month || row.code || row.account || row.amount) : [];
}

function getDraftPayload() {
  return {
    version: draftStorage.version,
    savedAt: new Date().toISOString(),
    importMode: state.importMode,
    ruleScope: state.ruleScope,
    forecastStart: state.forecastStart,
    forecastMonths: state.forecastMonths,
    goLiveDate: state.goLiveDate,
    activeTab: state.activeTab,
    filters: cloneData(state.filters),
    actualRows: cloneData(state.actualRows),
    actualImportFlags: cloneData(state.actualImportFlags),
    flowRows: cloneData(state.flowRows),
    flowImportName: state.flowImportName,
    flowImportRowCount: state.flowImportRowCount,
    flowImportMatchedChannels: cloneData(state.flowImportMatchedChannels),
    flowImportUnmatchedChannels: cloneData(state.flowImportUnmatchedChannels),
    flowImportSourceFormat: state.flowImportSourceFormat,
    historicalBreakdownMonth: state.historicalBreakdownMonth,
    historicalBreakdownStartMonth: state.historicalBreakdownStartMonth,
    historicalBreakdownEndMonth: state.historicalBreakdownEndMonth,
    historicalBreakdownMode: state.historicalBreakdownMode,
    historicalBurnWindow: state.historicalBurnWindow,
    projectBudgetAmount: state.projectBudgetAmount,
    pivotExpanded: state.pivotExpanded,
    budgetAlertLaunchMonth: state.budgetAlertLaunchMonth,
    budgetAlertMarketingBudget: state.budgetAlertMarketingBudget,
    budgetAlertHcGrowthRate: state.budgetAlertHcGrowthRate,
    budgetAlertExpanded: state.budgetAlertExpanded,
    delayMonths: state.delayMonths,
    delayHcGrowthRate: state.delayHcGrowthRate,
    paybackMonths: state.paybackMonths,
    grossMarginRate: state.grossMarginRate,
    resultRevenueView: state.resultRevenueView,
    stressRevenueDownRate: state.stressRevenueDownRate,
    stressUaUpRate: state.stressUaUpRate,
    revenueAggregateOverrides: {
      rawRevenueValues: cloneData(state.revenueAggregateOverrides.rawRevenueValues),
      rawVatRateValues: cloneData(state.revenueAggregateOverrides.rawVatRateValues),
      rawPlatformRateValues: cloneData(state.revenueAggregateOverrides.rawPlatformRateValues)
    },
    hcForecastRows: state.hcForecastRows.map((row) => ({
      code: row.code,
      feeType: row.feeType,
      estimateMethodRaw: row.estimateMethodRaw,
      costDriver: row.costDriver,
      suggestedLogic: row.suggestedLogic,
      referencePeriodRaw: row.referencePeriodRaw,
      forecastParamRaw: row.forecastParamRaw,
      forecastLogicRaw: row.forecastLogicRaw,
      rawValues: cloneData(row.rawValues || {})
    })),
    revenueForecastRows: state.revenueForecastRows.map((row) => ({
      code: row.code,
      feeType: row.feeType,
      estimateMethodRaw: row.estimateMethodRaw,
      costDriver: row.costDriver,
      suggestedLogic: row.suggestedLogic,
      referencePeriodRaw: row.referencePeriodRaw,
      forecastParamRaw: row.forecastParamRaw,
      forecastLogicRaw: row.forecastLogicRaw,
      rawValues: cloneData(row.rawValues || {})
    })),
    costForecastRows: state.costForecastRows.map((row) => ({
      code: row.code,
      feeType: row.feeType,
      estimateMethodRaw: row.estimateMethodRaw,
      costDriver: row.costDriver,
      suggestedLogic: row.suggestedLogic,
      referencePeriodRaw: row.referencePeriodRaw,
      forecastParamRaw: row.forecastParamRaw,
      forecastLogicRaw: row.forecastLogicRaw,
      rawValues: cloneData(row.rawValues || {})
    })),
    marketingForecastRows: state.marketingForecastRows.map((row) => ({
      code: row.code,
      feeType: row.feeType,
      estimateMethodRaw: row.estimateMethodRaw,
      costDriver: row.costDriver,
      suggestedLogic: row.suggestedLogic,
      referencePeriodRaw: row.referencePeriodRaw,
      forecastParamRaw: row.forecastParamRaw,
      forecastLogicRaw: row.forecastLogicRaw,
      rawValues: cloneData(row.rawValues || {})
    })),
    rules: cloneData(state.rules),
    templateLayoutRows: cloneData(state.templateLayoutRows),
    postPivot: cloneData(state.postPivot),
    importMetaText: {
      post: els.pnlPostMeta?.textContent || "",
      pre: els.pnlPreMeta?.textContent || "",
      allocated: els.pnlAllocMeta?.textContent || "",
      flow: els.flowMeta?.textContent || ""
    }
  };
}

function getComparableDraftPayload() {
  const payload = getDraftPayload();
  delete payload.savedAt;
  return payload;
}

function getDraftSignature() {
  return JSON.stringify(getComparableDraftPayload());
}

function normalizeDraftName(raw) {
  const text = String(raw || "").trim().replace(/\s+/g, " ");
  return text.slice(0, 40);
}

function getDefaultDraftName() {
  const projectName = String(state.postPivot.displayName || "项目测算").trim() || "项目测算";
  return `${projectName}-方案`;
}

function createDraftId() {
  return `draft-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function getDraftEntryById(draftId = state.persistence.activeDraftId) {
  return state.persistence.drafts.find((draft) => draft.id === draftId) || null;
}

function getSelectedDraftId() {
  return els.draftSlotSelect?.value || state.persistence.activeDraftId || "";
}

function refreshDraftControls() {
  if (els.draftSlotSelect) {
    const preferredSelectedId = els.draftSlotSelect.value || state.persistence.activeDraftId || state.persistence.drafts[0]?.id || "";
    els.draftSlotSelect.innerHTML = state.persistence.drafts.length
      ? state.persistence.drafts.map((draft) => `<option value="${draft.id}">${draft.name} · ${formatTimestamp(draft.savedAt) || "未保存时间"}</option>`).join("")
      : `<option value="">暂无本地草稿</option>`;
    els.draftSlotSelect.value = state.persistence.drafts.some((draft) => draft.id === preferredSelectedId) ? preferredSelectedId : (state.persistence.drafts[0]?.id || "");
    els.draftSlotSelect.disabled = !state.persistence.available || !state.persistence.drafts.length;
  }

  const selectedDraft = getDraftEntryById(getSelectedDraftId());
  const nextName = state.persistence.activeDraftName || selectedDraft?.name || getDefaultDraftName();
  state.persistence.activeDraftName = normalizeDraftName(nextName) || getDefaultDraftName();
  if (!state.persistence.drafts.some((draft) => draft.id === state.persistence.activeDraftId)) {
    state.persistence.activeDraftId = "";
  }

  if (els.draftSlotNameInput) {
    els.draftSlotNameInput.value = state.persistence.activeDraftName;
    els.draftSlotNameInput.disabled = !state.persistence.available;
  }
  if (els.draftSaveAsButton) {
    els.draftSaveAsButton.disabled = !state.persistence.available;
  }
}

function setDraftStatus(text, { lastSavedAt = state.persistence.lastSavedAt, hasDraft = state.persistence.hasDraft } = {}) {
  state.persistence.statusText = text;
  state.persistence.lastSavedAt = lastSavedAt;
  state.persistence.hasDraft = hasDraft;
  if (els.draftStatusText) {
    els.draftStatusText.textContent = text;
  }
  refreshDraftControls();
  if (els.draftRestoreButton) {
    els.draftRestoreButton.disabled = !state.persistence.available || !hasDraft;
  }
  if (els.draftClearButton) {
    els.draftClearButton.disabled = !state.persistence.available || !hasDraft;
  }
  if (els.draftSaveButton) {
    els.draftSaveButton.disabled = !state.persistence.available;
  }
}

function buildDraftCollectionPayload() {
  return {
    version: draftStorage.version,
    activeDraftId: state.persistence.activeDraftId,
    drafts: state.persistence.drafts.map((draft) => ({
      id: draft.id,
      name: draft.name,
      savedAt: draft.savedAt,
      payload: draft.payload
    }))
  };
}

function getLegacyDraftPayload(storage) {
  try {
    const raw = storage.getItem(draftStorage.legacyKey);
    if (!raw) return null;
    const payload = JSON.parse(raw);
    if (!payload || Number(payload.version) !== 1) return null;
    return payload;
  } catch {
    return null;
  }
}

function readDraftCollection() {
  const storage = safeLocalStorage();
  if (!storage) return null;
  try {
    const raw = storage.getItem(draftStorage.key);
    if (raw) {
      const payload = JSON.parse(raw);
      if (Number(payload?.version) === draftStorage.version && Array.isArray(payload?.drafts)) {
        return {
          activeDraftId: String(payload.activeDraftId || ""),
          drafts: payload.drafts.map((draft) => ({
            id: String(draft.id || createDraftId()),
            name: normalizeDraftName(draft.name) || "未命名草稿",
            savedAt: String(draft.savedAt || ""),
            payload: draft.payload
          }))
        };
      }
    }

    const legacyPayload = getLegacyDraftPayload(storage);
    if (!legacyPayload) return null;
    const migratedDraft = {
      id: createDraftId(),
      name: normalizeDraftName(legacyPayload.postPivot?.displayName || "历史草稿") || "历史草稿",
      savedAt: String(legacyPayload.savedAt || new Date().toISOString()),
      payload: legacyPayload
    };
    const collection = { activeDraftId: migratedDraft.id, drafts: [migratedDraft] };
    storage.setItem(draftStorage.key, JSON.stringify({ version: draftStorage.version, ...collection }));
    storage.removeItem(draftStorage.legacyKey);
    return collection;
  } catch {
    return null;
  }
}

function persistDraft(reason = "auto", { forceNew = false } = {}) {
  if (!state.persistence.available || suspendAutoSave) return false;
  const storage = safeLocalStorage();
  if (!storage) {
    state.persistence.available = false;
    setDraftStatus("当前浏览器无法使用本地草稿保存", { hasDraft: false, lastSavedAt: "" });
    return false;
  }

  try {
    const payload = getDraftPayload();
    const signature = getDraftSignature();
    const draftName = normalizeDraftName(els.draftSlotNameInput?.value || state.persistence.activeDraftName || getDefaultDraftName()) || getDefaultDraftName();
    const draftId = forceNew || !state.persistence.activeDraftId ? createDraftId() : state.persistence.activeDraftId;
    const nextDraft = {
      id: draftId,
      name: draftName,
      savedAt: payload.savedAt,
      payload
    };
    const remainingDrafts = state.persistence.drafts.filter((draft) => draft.id !== draftId);
    state.persistence.drafts = [nextDraft, ...remainingDrafts].sort((a, b) => String(b.savedAt || "").localeCompare(String(a.savedAt || "")));
    state.persistence.activeDraftId = draftId;
    state.persistence.activeDraftName = draftName;
    storage.setItem(draftStorage.key, JSON.stringify({ version: draftStorage.version, activeDraftId: draftId, drafts: state.persistence.drafts }));
    const lastSavedAt = nextDraft.savedAt;
    const actionLabel = reason === "manual" ? "已手动保存草稿" : "已自动保存草稿";
    state.persistence.lastSavedSignature = signature;
    state.persistence.dirty = false;
    setDraftStatus(`${actionLabel}到「${draftName}」：${formatTimestamp(lastSavedAt)}`, { lastSavedAt, hasDraft: true });
    return true;
  } catch {
    setDraftStatus("草稿保存失败，请检查浏览器存储空间", { hasDraft: state.persistence.hasDraft });
    return false;
  }
}

function scheduleAutoSave() {
  if (!state.persistence.available || suspendAutoSave) return;
  const currentSignature = getDraftSignature();
  state.persistence.dirty = currentSignature !== state.persistence.lastSavedSignature;
  if (!state.persistence.dirty) {
    window.clearTimeout(autoSaveTimer);
    if (state.persistence.hasDraft && state.persistence.lastSavedAt) {
      const draftName = getDraftEntryById(getSelectedDraftId())?.name || state.persistence.activeDraftName || "当前草稿";
      setDraftStatus(`当前内容已保存到「${draftName}」：${formatTimestamp(state.persistence.lastSavedAt)}`, {
        hasDraft: true,
        lastSavedAt: state.persistence.lastSavedAt
      });
    } else if (!state.persistence.hasDraft) {
      setDraftStatus("当前内容尚未保存到任何草稿槽位", { hasDraft: false, lastSavedAt: "" });
    }
    return;
  }
  window.clearTimeout(autoSaveTimer);
  const draftName = normalizeDraftName(els.draftSlotNameInput?.value || state.persistence.activeDraftName || getDefaultDraftName()) || getDefaultDraftName();
  setDraftStatus(`检测到未保存改动，正在准备保存到「${draftName}」...`, { hasDraft: state.persistence.hasDraft });
  autoSaveTimer = window.setTimeout(() => {
    persistDraft("auto");
  }, draftStorage.autoSaveDelay);
}

function applyImportedMetaText(metaText = {}) {
  if (typeof metaText.post === "string" && els.pnlPostMeta) els.pnlPostMeta.textContent = metaText.post;
  if (typeof metaText.pre === "string" && els.pnlPreMeta) els.pnlPreMeta.textContent = metaText.pre;
  if (typeof metaText.allocated === "string" && els.pnlAllocMeta) els.pnlAllocMeta.textContent = metaText.allocated;
  if (typeof metaText.flow === "string" && els.flowMeta) els.flowMeta.textContent = metaText.flow;
}

function restoreDraftPayload(payload) {
  if (!payload || Number(payload.version) !== draftStorage.version) return false;

  const savedImportMode = ["post", "split"].includes(payload.importMode) ? payload.importMode : "post";
  const savedRuleScope = ["post", "pre", "allocated"].includes(payload.ruleScope) ? payload.ruleScope : "post";
  const savedForecastStart = normalizeMonth(payload.forecastStart);
  const savedForecastMonths = Math.min(Math.max(Number(payload.forecastMonths) || 12, 3), 36);

  suspendAutoSave = true;

  state.importMode = savedImportMode;
  state.ruleScope = savedImportMode === "post" ? "post" : savedRuleScope;
  state.forecastStart = savedForecastStart || state.forecastStart;
  state.forecastMonths = savedForecastMonths;
  state.goLiveDate = /^\d{4}-\d{2}-\d{2}$/.test(String(payload.goLiveDate || "")) ? String(payload.goLiveDate) : "";
  state.activeTab = ["inputs", "rules", "results"].includes(payload.activeTab) ? payload.activeTab : "inputs";
  state.filters = {
    source: String(payload.filters?.source || "all"),
    type: String(payload.filters?.type || "all"),
    search: String(payload.filters?.search || "")
  };
  state.actualRows = {
    post: normalizeActualRowsSnapshot(payload.actualRows?.post),
    pre: normalizeActualRowsSnapshot(payload.actualRows?.pre),
    allocated: normalizeActualRowsSnapshot(payload.actualRows?.allocated)
  };
  state.actualImportFlags = {
    post: Boolean(payload.actualImportFlags?.post),
    pre: Boolean(payload.actualImportFlags?.pre),
    allocated: Boolean(payload.actualImportFlags?.allocated)
  };
  state.flowRows = normalizeFlowRows(payload.flowRows || []);
  state.flowImportName = String(payload.flowImportName || "");
  state.flowImportRowCount = Number(payload.flowImportRowCount || state.flowRows.length || 0);
  state.flowImportMatchedChannels = Array.isArray(payload.flowImportMatchedChannels) ? payload.flowImportMatchedChannels.map((item) => String(item)) : [];
  state.flowImportUnmatchedChannels = Array.isArray(payload.flowImportUnmatchedChannels) ? payload.flowImportUnmatchedChannels.map((item) => String(item)) : [];
  state.flowImportSourceFormat = String(payload.flowImportSourceFormat || "");
  state.historicalBreakdownMonth = normalizeMonth(payload.historicalBreakdownMonth || state.historicalBreakdownMonth);
  state.historicalBreakdownStartMonth = normalizeMonth(payload.historicalBreakdownStartMonth || payload.historicalBreakdownMonth || state.historicalBreakdownStartMonth);
  state.historicalBreakdownEndMonth = normalizeMonth(payload.historicalBreakdownEndMonth || payload.historicalBreakdownMonth || state.historicalBreakdownEndMonth);
  state.historicalBreakdownMode = ["cumulative", "average", "single"].includes(String(payload.historicalBreakdownMode || ""))
    ? String(payload.historicalBreakdownMode)
    : state.historicalBreakdownMode;
  state.historicalBurnWindow = [1, 2, 3].includes(Number(payload.historicalBurnWindow)) ? Number(payload.historicalBurnWindow) : state.historicalBurnWindow;
  state.projectBudgetAmount = Math.max(Number(payload.projectBudgetAmount) || 0, 0);
  state.pivotExpanded = Boolean(payload.pivotExpanded);
  state.budgetAlertLaunchMonth = normalizeMonth(payload.budgetAlertLaunchMonth || state.budgetAlertLaunchMonth);
  state.budgetAlertMarketingBudget = Math.max(Number(payload.budgetAlertMarketingBudget) || 0, 0);
  state.budgetAlertHcGrowthRate = Math.min(Math.max(Number(payload.budgetAlertHcGrowthRate) || 0, 0), 1);
  state.budgetAlertExpanded = Boolean(payload.budgetAlertExpanded);
  state.delayMonths = Math.min(
    Math.max(Number.isFinite(Number(payload.delayMonths)) ? Number(payload.delayMonths) : state.delayMonths, 0),
    12
  );
  state.delayHcGrowthRate = Math.min(Math.max(Number(payload.delayHcGrowthRate) || 0, 0), 1);
  state.paybackMonths = Math.min(Math.max(Number(payload.paybackMonths) || state.paybackMonths, 1), 24);
  state.grossMarginRate = Math.min(Math.max(Number(payload.grossMarginRate) || state.grossMarginRate, 0.1), 0.9);
  state.resultRevenueView = ["region", "platform"].includes(String(payload.resultRevenueView || "")) ? String(payload.resultRevenueView) : "region";
  state.stressRevenueDownRate = Math.min(Math.max(Number(payload.stressRevenueDownRate) || 0, 0), 0.6);
  state.stressUaUpRate = Math.min(Math.max(Number(payload.stressUaUpRate) || 0, 0), 1);
  state.templateLayoutRows = normalizeTemplateLayoutRows(
    Array.isArray(payload.templateLayoutRows) && payload.templateLayoutRows.length
      ? payload.templateLayoutRows
      : state.templateLayoutRows
  );
  state.postPivot = payload.postPivot && Array.isArray(payload.postPivot.rows)
    ? cloneData(payload.postPivot)
    : state.postPivot;

  rebuildHcForecastRows();
  const savedHcRowMap = new Map((payload.hcForecastRows || []).map((row) => [String(row.code || ""), row]));
  state.hcForecastRows.forEach((row) => {
    const savedHcRow = savedHcRowMap.get(row.code);
    if (!savedHcRow) return;
    ["feeType", "estimateMethodRaw", "costDriver", "suggestedLogic", "referencePeriodRaw", "forecastParamRaw", "forecastLogicRaw"].forEach((field) => {
      if (savedHcRow[field] !== undefined) row[field] = String(savedHcRow[field] || "");
    });
    row.suggestedValueText = deriveHcSuggestedValueText(row.code, row.referencePeriodRaw);
    Object.keys(row.rawValues).forEach((month) => {
      if (savedHcRow.rawValues?.[month] !== undefined) {
        row.rawValues[month] = String(savedHcRow.rawValues[month]);
      }
    });
  });
  recomputeHcForecastValues();

  rebuildRevenueForecastRows();
  const savedRevenueRowMap = new Map((payload.revenueForecastRows || []).map((row) => [String(row.code || ""), row]));
  state.revenueForecastRows.forEach((row) => {
    const savedRow = savedRevenueRowMap.get(row.code);
    if (!savedRow) return;
    ["feeType", "estimateMethodRaw", "costDriver", "suggestedLogic", "referencePeriodRaw", "forecastParamRaw", "forecastLogicRaw"].forEach((field) => {
      if (savedRow[field] !== undefined) row[field] = String(savedRow[field] || "");
    });
    row.suggestedValueText = deriveRevenueSuggestedValueText(row.code, row.referencePeriodRaw);
    Object.keys(row.rawValues).forEach((month) => {
      if (savedRow.rawValues?.[month] !== undefined) {
        row.rawValues[month] = String(savedRow.rawValues[month]);
      }
    });
  });
  recomputeRevenueForecastValues();

  rebuildCostForecastRows();
  const savedCostRowMap = new Map((payload.costForecastRows || []).map((row) => [String(row.code || ""), row]));
  state.costForecastRows.forEach((row) => {
    const savedRow = savedCostRowMap.get(row.code);
    if (!savedRow) return;
    ["feeType", "estimateMethodRaw", "costDriver", "suggestedLogic", "referencePeriodRaw", "forecastParamRaw", "forecastLogicRaw"].forEach((field) => {
      if (savedRow[field] !== undefined) row[field] = String(savedRow[field] || "");
    });
    row.suggestedValueText = deriveCostSuggestedValueText(row.code, row.referencePeriodRaw);
    Object.keys(row.rawValues).forEach((month) => {
      if (savedRow.rawValues?.[month] !== undefined) {
        row.rawValues[month] = String(savedRow.rawValues[month]);
      }
    });
  });
  recomputeCostForecastValues();

  rebuildMarketingForecastRows();
  const savedMarketingRowMap = new Map((payload.marketingForecastRows || []).map((row) => [String(row.code || ""), row]));
  state.marketingForecastRows.forEach((row) => {
    const savedRow = savedMarketingRowMap.get(row.code);
    if (!savedRow) return;
    ["feeType", "estimateMethodRaw", "costDriver", "suggestedLogic", "referencePeriodRaw", "forecastParamRaw", "forecastLogicRaw"].forEach((field) => {
      if (savedRow[field] !== undefined) row[field] = String(savedRow[field] || "");
    });
    row.suggestedValueText = deriveMarketingSuggestedValueText(row.code, row.referencePeriodRaw);
    Object.keys(row.rawValues).forEach((month) => {
      if (savedRow.rawValues?.[month] !== undefined) {
        row.rawValues[month] = String(savedRow.rawValues[month]);
      }
    });
  });
  recomputeMarketingForecastValues();

  state.rules = {
    post: Array.isArray(payload.rules?.post) ? cloneData(payload.rules.post) : [],
    pre: Array.isArray(payload.rules?.pre) ? cloneData(payload.rules.pre) : [],
    allocated: Array.isArray(payload.rules?.allocated) ? cloneData(payload.rules.allocated) : []
  };
  syncRuleBoundaries();
  refreshRules();
  syncInputs();
  applyImportedMetaText(payload.importMetaText);
  state.persistence.lastSavedSignature = getDraftSignature();
  state.persistence.dirty = false;
  suspendAutoSave = false;
  return true;
}

function restoreSelectedDraft() {
  const selectedDraft = getDraftEntryById(getSelectedDraftId());
  if (!selectedDraft?.payload) return false;
  state.persistence.activeDraftId = selectedDraft.id;
  state.persistence.activeDraftName = selectedDraft.name;
  const restored = restoreDraftPayload(selectedDraft.payload);
  if (restored) {
    setDraftStatus(`已恢复草稿「${selectedDraft.name}」：${formatTimestamp(selectedDraft.savedAt)}`, {
      hasDraft: true,
      lastSavedAt: selectedDraft.savedAt
    });
  }
  return restored;
}

function deleteSelectedDraft() {
  const storage = safeLocalStorage();
  if (!storage) return;
  const selectedDraftId = getSelectedDraftId();
  const selectedDraft = getDraftEntryById(selectedDraftId);
  if (!selectedDraft) return;
  const isDeletingActiveDraft = selectedDraftId === state.persistence.activeDraftId;
  state.persistence.drafts = state.persistence.drafts.filter((draft) => draft.id !== selectedDraftId);
  if (isDeletingActiveDraft) {
    state.persistence.activeDraftId = "";
    state.persistence.activeDraftName = getDefaultDraftName();
    state.persistence.lastSavedSignature = getDraftSignature();
    state.persistence.dirty = false;
  }
  storage.setItem(draftStorage.key, JSON.stringify(buildDraftCollectionPayload()));
  const nextDraft = state.persistence.drafts[0] || null;
  setDraftStatus(
    state.persistence.drafts.length
      ? `已删除草稿「${selectedDraft.name}」，其余草稿仍可继续恢复`
      : "已删除最后一个草稿槽位",
    { hasDraft: state.persistence.drafts.length > 0, lastSavedAt: nextDraft?.savedAt || "" }
  );
}

function escapeCsvValue(value) {
  const text = value === null || value === undefined ? "" : String(value);
  if (/["\r\n,]/.test(text)) return `"${text.replace(/"/g, "\"\"")}"`;
  return text;
}

function buildCsvContent(headers, rows) {
  const lines = [headers.map(escapeCsvValue).join(",")];
  rows.forEach((row) => {
    lines.push(row.map(escapeCsvValue).join(","));
  });
  return `\uFEFF${lines.join("\r\n")}`;
}

function triggerDownload(fileName, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getExportBaseName() {
  const rawName = String(state.postPivot.displayName || "项目损益测算").trim() || "项目损益测算";
  const safeName = rawName.replace(/[\\/:*?"<>|]/g, "-");
  const dateText = new Date().toISOString().slice(0, 10);
  return `${safeName}-V2.9-${dateText}`;
}

function exportSummaryCsv() {
  const model = calculateModel();
  const rows = model.monthly.map((row) => [
    row.month,
    row.grossRevenue,
    row.vatAmount,
    row.platformAmount,
    row.netRevenue,
    row.direct,
    row.allocated,
    row.merged,
    row.profit
  ]);
  const csv = buildCsvContent(
    ["月份", "A2 项目流水", "A3 增值税", "A4 平台费", "A5 净收入", "项目组直属", "分摊值", "合并成本", "项目利润"],
    rows
  );
  triggerDownload(`${getExportBaseName()}-月度汇总.csv`, csv, "text/csv;charset=utf-8");
}

function exportProfitBridgeCsv() {
  const model = calculateModel();
  const rows = model.profitBridge.map((row, index) => [
    row.month,
    row.a5,
    row.a97,
    row.a91,
    row.a9,
    row.a10,
    row.a99,
    row.a14,
    row.a18,
    row.a13,
    model.monthly[index]?.profit || 0
  ]);
  const csv = buildCsvContent(
    ["月份", "A5 净收入", "A97 其他收入", "A91 内采收入", "A9 毛利润", "A10 市场费用", "A99 发行利润", "A14 人力费用", "A18 外包费用", "A13 行政及其他", "项目利润"],
    rows
  );
  triggerDownload(`${getExportBaseName()}-利润链路.csv`, csv, "text/csv;charset=utf-8");
}

function exportSnapshotJson() {
  const model = calculateModel();
  const payload = {
    exportedAt: new Date().toISOString(),
    projectName: state.postPivot.displayName || "项目损益测算",
    forecastStart: state.forecastStart,
    forecastMonths: state.forecastMonths,
    importMode: state.importMode,
    ruleScope: getActiveScope(),
    metrics: {
      totalGrossRevenue: model.totalGrossRevenue,
      totalNetRevenue: model.totalNetRevenue,
      totalCost: model.totalCost,
      totalProfit: model.totalProfit,
      roi: model.totalCost ? model.totalProfit / model.totalCost : 0
    },
    monthly: model.monthly,
    profitBridge: model.profitBridge,
    bySource: model.bySource,
    assumptions: {
      hcForecastRows: state.hcForecastRows,
      revenueForecastRows: state.revenueForecastRows,
      costForecastRows: state.costForecastRows,
      marketingForecastRows: state.marketingForecastRows,
      rules: state.rules
    }
  };
  triggerDownload(`${getExportBaseName()}-测算包.json`, JSON.stringify(payload, null, 2), "application/json;charset=utf-8");
}

function buildReportSummaryChartSvg(model) {
  const maxValue = Math.max(...model.monthly.flatMap((row) => [row.netRevenue, row.merged, Math.abs(row.profit)]), 1);
  const barSlot = 760 / Math.max(model.monthly.length, 1);
  const bars = model.monthly.map((row, index) => {
    const x = 70 + index * barSlot;
    const revenueHeight = (row.netRevenue / maxValue) * 180;
    const costHeight = (row.merged / maxValue) * 180;
    const profitY = 250 - (row.profit / maxValue) * 160;
    return `
      <rect x="${x}" y="${250 - revenueHeight}" width="16" height="${revenueHeight}" fill="#0f6cbd" opacity="0.82"></rect>
      <rect x="${x + 22}" y="${250 - costHeight}" width="16" height="${costHeight}" fill="#f77f00" opacity="0.82"></rect>
      <circle cx="${x + 48}" cy="${profitY}" r="4" fill="${row.profit >= 0 ? "#1f8f6a" : "#c03b33"}"></circle>
      <text x="${x + 18}" y="282" font-size="11" fill="#64748b">${row.month.slice(5)}</text>
    `;
  }).join("");

  return `
    <svg class="report-chart" viewBox="0 0 920 320" preserveAspectRatio="none">
      <line x1="52" y1="250" x2="888" y2="250" stroke="rgba(29,42,58,0.18)"></line>
      <text x="62" y="28" font-size="12" fill="#0f6cbd">净收入</text>
      <text x="118" y="28" font-size="12" fill="#f77f00">合并成本</text>
      <text x="198" y="28" font-size="12" fill="#1f8f6a">项目利润</text>
      ${bars}
    </svg>
  `;
}

function buildReportProfitBridgeChartSvg(model) {
  const series = [
    { key: "a9", label: "A9 毛利润", color: "#0f6cbd" },
    { key: "a99", label: "A99 发行利润", color: "#f77f00" },
    { key: "profit", label: "项目利润", color: "#1f8f6a" }
  ];
  const maxValue = Math.max(
    ...model.profitBridge.flatMap((row, index) => [Math.abs(row.a9), Math.abs(row.a99), Math.abs(model.monthly[index]?.profit || 0)]),
    1
  );
  const baseY = 252;
  const scale = 145 / maxValue;
  const pointX = (index) => 70 + index * (760 / Math.max(model.profitBridge.length - 1, 1));
  const pointY = (value) => baseY - value * scale;
  const pathFor = (getter) => model.profitBridge.map((row, index) => `${index === 0 ? "M" : "L"} ${pointX(index)} ${pointY(getter(row, index))}`).join(" ");
  const legend = series.map((item, index) => `
    <circle cx="${70 + index * 150}" cy="28" r="5" fill="${item.color}"></circle>
    <text x="${82 + index * 150}" y="32" font-size="12" fill="${item.color}">${item.label}</text>
  `).join("");
  const labels = model.profitBridge.map((row, index) => `
    <text x="${pointX(index)}" y="286" font-size="11" text-anchor="middle" fill="#64748b">${row.month.slice(5)}</text>
  `).join("");
  const dots = series.map((item) => model.profitBridge.map((row, index) => {
    const value = item.key === "profit" ? (model.monthly[index]?.profit || 0) : row[item.key];
    return `<circle cx="${pointX(index)}" cy="${pointY(value)}" r="4" fill="${item.color}"></circle>`;
  }).join("")).join("");

  return `
    <svg class="report-chart" viewBox="0 0 920 320" preserveAspectRatio="none">
      <line x1="52" y1="${baseY}" x2="888" y2="${baseY}" stroke="rgba(29,42,58,0.18)"></line>
      <line x1="52" y1="98" x2="888" y2="98" stroke="rgba(29,42,58,0.08)" stroke-dasharray="4 4"></line>
      ${legend}
      <path d="${pathFor((row) => row.a9)}" fill="none" stroke="#0f6cbd" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
      <path d="${pathFor((row) => row.a99)}" fill="none" stroke="#f77f00" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
      <path d="${pathFor((row, index) => model.monthly[index]?.profit || 0)}" fill="none" stroke="#1f8f6a" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
      ${dots}
      ${labels}
    </svg>
  `;
}

function exportShareReportHtml() {
  const model = calculateModel();
  const roi = model.totalCost ? model.totalProfit / model.totalCost : 0;
  const projectName = state.postPivot.displayName || "项目损益测算";
  const total = Math.max(...Object.values(model.bySource), 1);
  const metrics = [
    ["累计项目流水(A2)", compactMoney(model.totalGrossRevenue), ""],
    ["累计净收入(A5)", compactMoney(model.totalNetRevenue), ""],
    ["累计合并成本", compactMoney(model.totalCost), ""],
    ["累计项目利润", compactMoney(model.totalProfit), model.totalProfit >= 0 ? "positive" : "negative"],
    ["ROI", ratio(roi), roi >= 0 ? "positive" : "negative"],
    ["盈利月份数", `${model.monthly.filter((row) => row.profit >= 0).length}/${model.monthly.length}`, ""],
    ["预测起始月", state.forecastStart, ""],
    ["预测月数", `${state.forecastMonths} 个月`, ""]
  ];
  const metricHtml = metrics.map(([label, value, tone]) => `
    <article class="metric">
      <div class="label">${label}</div>
      <div class="value ${tone}">${value}</div>
    </article>
  `).join("");
  const sourceHtml = Object.entries(model.bySource).map(([label, value]) => `
    <div class="bar-row">
      <strong>${label}</strong>
      <div class="bar-track"><div class="bar-fill" style="width:${(value / total) * 100}%"></div></div>
      <span>${compactMoney(value)}</span>
    </div>
  `).join("");
  const summaryRows = model.monthly.map((row) => `
    <tr>
      <td>${row.month}</td>
      <td>${compactMoney(row.grossRevenue)}</td>
      <td>${compactMoney(row.vatAmount)}</td>
      <td>${compactMoney(row.platformAmount)}</td>
      <td>${compactMoney(row.netRevenue)}</td>
      <td>${compactMoney(row.merged)}</td>
      <td class="${row.profit >= 0 ? "positive" : "negative"}">${compactMoney(row.profit)}</td>
    </tr>
  `).join("");
  const bridgeRows = model.profitBridge.map((row, index) => `
    <tr>
      <td>${row.month}</td>
      <td>${compactMoney(row.a5)}</td>
      <td>${compactMoney(row.a97)}</td>
      <td>${compactMoney(row.a91)}</td>
      <td>${compactMoney(row.a9)}</td>
      <td>${compactMoney(row.a10)}</td>
      <td>${compactMoney(row.a99)}</td>
      <td>${compactMoney(model.monthly[index]?.profit || 0)}</td>
    </tr>
  `).join("");
  const summaryChart = buildReportSummaryChartSvg(model);
  const bridgeChart = buildReportProfitBridgeChartSvg(model);

  const html = `<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${projectName} - 损益测算分享报告</title>
    <style>${reportStyles}</style>
  </head>
  <body>
    <main class="report-shell">
      <section class="hero">
        <p class="eyebrow">Game P&L Snapshot</p>
        <h1>${projectName} 损益测算分享报告</h1>
        <p class="subtitle">该报告由 V2.9 测算台自动导出，适合用于内部同步当前测算结论、成本结构和利润链路。</p>
        <div class="meta">
          <span class="chip">导出口径：${state.importMode === "post" ? "分摊后" : "分摊前 + 分摊值"}</span>
          <span class="chip">当前规则视图：${getActiveScope() === "post" ? "分摊后" : getActiveScope() === "pre" ? "分摊前" : "分摊值"}</span>
          <span class="chip">预测区间：${state.forecastStart} 起，共 ${state.forecastMonths} 个月</span>
          <span class="chip">导出时间：${new Date().toLocaleString("zh-CN")}</span>
        </div>
      </section>
      <section class="metric-grid">${metricHtml}</section>
      <section class="chart-grid">
        <section class="panel">
          <div class="section-head">
            <h2>月度走势</h2>
            <p>用柱状和点位快速查看净收入、合并成本与项目利润的月度变化。</p>
          </div>
          ${summaryChart}
        </section>
        <section class="panel">
          <div class="section-head">
            <h2>利润链路图</h2>
            <p>用折线对比毛利润、发行利润与最终项目利润的变化趋势。</p>
          </div>
          ${bridgeChart}
        </section>
      </section>
      <section class="split-grid">
        <section class="panel">
          <div class="section-head">
            <h2>月度汇总</h2>
            <p>按月展示 A2、A3、A4、A5、合并成本与项目利润，适合快速复核关键波动。</p>
          </div>
          <table>
            <thead>
              <tr><th>月份</th><th>A2 项目流水</th><th>A3 增值税</th><th>A4 平台费</th><th>A5 净收入</th><th>合并成本</th><th>项目利润</th></tr>
            </thead>
            <tbody>${summaryRows}</tbody>
          </table>
        </section>
        <section class="panel">
          <div class="section-head">
            <h2>费用来源拆分</h2>
            <p>区分项目组直属与支持部门分摊值，便于在对外同步时快速说明成本结构。</p>
          </div>
          <div class="bar-list">${sourceHtml}</div>
        </section>
      </section>
      <section class="panel">
        <div class="section-head">
          <h2>利润链路</h2>
          <p>展示净收入、毛利润、发行利润到最终项目利润的关键传导路径。</p>
        </div>
        <table>
          <thead>
            <tr><th>月份</th><th>A5 净收入</th><th>A97 其他收入</th><th>A91 内采收入</th><th>A9 毛利润</th><th>A10 市场费用</th><th>A99 发行利润</th><th>项目利润</th></tr>
          </thead>
          <tbody>${bridgeRows}</tbody>
        </table>
      </section>
      <section class="panel footnote">
        本报告为静态 HTML 快照，便于转发和归档；若需继续编辑测算假设，请回到 V2.9 测算台原页面。
      </section>
    </main>
  </body>
</html>`;

  triggerDownload(`${getExportBaseName()}-分享报告.html`, html, "text/html;charset=utf-8");
}

function normalizeHcRawInput(raw) {
  const text = String(raw ?? "").trim();
  if (!text) return "";
  return text;
}

function getHcMonths() {
  return monthRange(state.forecastStart, state.forecastMonths);
}

function getHcBlueprint(code) {
  return hcRowBlueprints[code] || {
    feeType: "变动费用-人力相关",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  };
}

function getHcHistorySummary(code) {
  const sourceRows = getHcSourceRows().filter((row) => row.code === code);
  const monthlyTotals = [...new Map(
    sourceRows.map((row) => [normalizeMonth(row.month || row.dt), Number(row.amount || 0)])
  ).entries()]
    .filter(([month]) => month)
    .sort((a, b) => a[0].localeCompare(b[0]));

  const latestValue = monthlyTotals.at(-1)?.[1] ?? 0;
  const recentAverage = avg(monthlyTotals.slice(-3).map(([, value]) => value));
  return {
    monthlyTotals,
    latestValue,
    recentAverage,
    latestText: monthlyTotals.length ? integerWithThousands(latestValue) : "-",
    averageText: monthlyTotals.length ? integerWithThousands(recentAverage) : "-"
  };
}

function getHcReferencePeriodMonths(referencePeriodRaw) {
  const text = String(referencePeriodRaw || "").trim();
  if (text === "近1个月") return 1;
  if (text === "近3个月") return 3;
  if (text === "近6个月") return 6;
  if (text === "近12个月") return 12;
  return 0;
}

function deriveHcSuggestedValueText(code, referencePeriodRaw) {
  const historySummary = getHcHistorySummary(code);
  const periodMonths = getHcReferencePeriodMonths(referencePeriodRaw);
  if (!periodMonths || !historySummary.monthlyTotals.length) return "-";
  const values = historySummary.monthlyTotals.slice(-periodMonths).map(([, value]) => Number(value || 0));
  if (!values.length) return "-";
  return integerWithThousands(avg(values));
}

function deriveHcForecastLogic(methodRaw, fallback = "") {
  const method = String(methodRaw || "").trim();
  if (!method) return fallback;
  if (method === "按实际需求输入或读取excel") return "按实际需求输入或读取excel";
  if (method === "费用率" || method === "人均成本") return "=预估参数*成本动因";
  if (method === "月费用额") return "=预估参数";
  if (method.startsWith("=")) return "按预估方法公式自动计算";
  return fallback;
}

function getHcRowByCode(code) {
  return state.hcForecastRows.find((item) => item.code === code) || null;
}

function getHcCellRawValue(code, month) {
  const row = getHcRowByCode(code);
  return row?.rawValues?.[month] ?? "";
}

function getHcCellNumericValue(code, month) {
  const row = getHcRowByCode(code);
  return Number(row?.values?.[month] || 0);
}

function getHcCellAddress(rowIndex, colIndex) {
  const code = hcAccountCodes[rowIndex] || "-";
  const month = formatMonthSlashLabel(getHcMonths()[colIndex]) || "-";
  return `${code} · ${month}`;
}

function formatMonthSlashLabel(month) {
  if (!month) return "";
  const normalized = normalizeMonth(month);
  if (!normalized) return String(month);
  const [year, monthText] = normalized.split("-");
  return `${year}/${Number(monthText)}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderHcMetaDatalist(id, options) {
  return `
    <datalist id="${id}">
      ${options.map((option) => `<option value="${escapeHtml(String(option || "").replace(/\s*\n\s*/g, " "))}"></option>`).join("")}
    </datalist>
  `;
}

function colIndexToLetters(colIndex) {
  let value = Number(colIndex) + 1;
  let text = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    text = String.fromCharCode(65 + remainder) + text;
    value = Math.floor((value - 1) / 26);
  }
  return text;
}

function lettersToColIndex(text) {
  return String(text || "").toUpperCase().split("").reduce((total, char) => total * 26 + (char.charCodeAt(0) - 64), 0) - 1;
}

function sanitizeFormulaExpression(expression) {
  return expression.replace(/\s+/g, "");
}

function normalizeHcFormulaMonthToken(token, fallbackMonth = "") {
  const text = String(token || "").trim();
  if (!text) return normalizeMonth(fallbackMonth);
  if (/^\d{4}[-/]\d{1,2}$/.test(text)) {
    const normalizedText = text.replace("/", "-");
    return normalizeMonth(normalizedText);
  }
  return normalizeMonth(text || fallbackMonth);
}

function buildHcFormulaReference(code, month) {
  return `${code}@${formatMonthSlashLabel(month)}`;
}

function parseHcFormulaReference(reference, fallbackMonth) {
  const [codePart, monthPart] = String(reference || "").split("@");
  return {
    code: String(codePart || "").trim(),
    month: normalizeHcFormulaMonthToken(monthPart, fallbackMonth)
  };
}

function insertTextAtSelection(input, text) {
  const start = Number.isInteger(input.selectionStart) ? input.selectionStart : input.value.length;
  const end = Number.isInteger(input.selectionEnd) ? input.selectionEnd : input.value.length;
  const nextValue = `${input.value.slice(0, start)}${text}${input.value.slice(end)}`;
  input.value = nextValue;
  const nextCaret = start + text.length;
  if (typeof input.setSelectionRange === "function") input.setSelectionRange(nextCaret, nextCaret);
  return nextValue;
}

function getActiveHcFormulaEditor() {
  const activeElement = document.activeElement;
  if (!activeElement) return null;

  if (activeElement.id === "hc-formula-input") {
    const activeCell = getActiveHcCellMeta();
    if (!activeCell.code || !activeCell.month) return null;
    return {
      element: activeElement,
      kind: "formulaBar",
      code: activeCell.code,
      month: activeCell.month,
      rowIndex: activeCell.rowIndex,
      colIndex: activeCell.colIndex
    };
  }

  const cellInput = activeElement.closest?.("input[data-hc-code][data-hc-month]");
  if (!cellInput) return null;
  return {
    element: cellInput,
    kind: "cellInput",
    code: cellInput.dataset.hcCode,
    month: cellInput.dataset.hcMonth,
    rowIndex: Number(cellInput.dataset.hcRow),
    colIndex: Number(cellInput.dataset.hcCol)
  };
}

function syncHcFormulaEditorValue(editor, nextValue) {
  updateHcValue(editor.code, editor.month, nextValue, { recompute: false });
  const formulaBar = document.getElementById("hc-formula-input");
  if (formulaBar && editor.kind !== "formulaBar") formulaBar.value = nextValue;
  if (editor.kind === "formulaBar") {
    const sourceInput = els.hcModule.querySelector(`input[data-hc-row="${editor.rowIndex}"][data-hc-col="${editor.colIndex}"]`);
    if (sourceInput) sourceInput.value = nextValue;
  }
}

function tryInsertHcFormulaReferenceFromCell(cell) {
  const editor = getActiveHcFormulaEditor();
  if (!editor) return false;
  if (!String(editor.element.value || "").trim().startsWith("=")) return false;

  const rowIndex = Number(cell.dataset.hcRow);
  const colIndex = Number(cell.dataset.hcCol);
  const code = hcAccountCodes[rowIndex];
  const month = getHcMonths()[colIndex];
  if (!code || !month) return false;

  const reference = buildHcFormulaReference(code, month);
  const nextValue = insertTextAtSelection(editor.element, reference);
  syncHcFormulaEditorValue(editor, nextValue);
  state.hcFormulaInsertLock = true;
  return true;
}

function evaluateHcFormula(formula, currentRowIndex, currentColIndex, visiting, cache) {
  const month = getHcMonths()[currentColIndex];
  const expression = sanitizeFormulaExpression(formula.slice(1))
    .replace(/A\d+(?:\.\d+)*(?:@\d{4}[/-]\d{1,2})?/g, (matchedReference) => {
      const { code, month: refMonth } = parseHcFormulaReference(matchedReference, month);
      const refRowIndex = hcAccountCodes.indexOf(code);
      const refColIndex = getHcMonths().indexOf(refMonth);
      if (refRowIndex < 0 || refColIndex < 0 || !refMonth) return "0";
      return String(evaluateHcCell(code, refMonth, refRowIndex, refColIndex, visiting, cache));
    });

  if (/[^0-9+\-*/().]/.test(expression)) return 0;
  try {
    const value = Function(`"use strict"; return (${expression || 0});`)();
    return Number.isFinite(Number(value)) ? Number(value) : 0;
  } catch {
    return 0;
  }
}

function evaluateHcCell(code, month, rowIndex, colIndex, visiting = new Set(), cache = new Map()) {
  const key = `${code}__${month}`;
  if (cache.has(key)) return cache.get(key);
  if (visiting.has(key)) return 0;
  visiting.add(key);
  const raw = getHcCellRawValue(code, month);
  let value = 0;
  if (raw.startsWith("=")) {
    value = evaluateHcFormula(raw, rowIndex, colIndex, visiting, cache);
  } else {
    value = parseNumericInput(raw);
  }
  visiting.delete(key);
  cache.set(key, value);
  return value;
}

function recomputeHcForecastValues() {
  const months = getHcMonths();
  const cache = new Map();
  state.hcForecastRows.forEach((row, rowIndex) => {
    row.values = row.values || {};
    months.forEach((month, colIndex) => {
      row.values[month] = evaluateHcCell(row.code, month, rowIndex, colIndex, new Set(), cache);
    });
  });
}

function setHcCellRawValue(code, month, rawValue, { recompute = true } = {}) {
  const row = getHcRowByCode(code);
  if (!row) return;
  row.rawValues = row.rawValues || {};
  row.rawValues[month] = normalizeHcRawInput(rawValue);
  if (recompute) recomputeHcForecastValues();
}

function normalizeHcSelection(selection = state.hcSelection) {
  return {
    startRow: Math.min(selection.startRow, selection.endRow),
    endRow: Math.max(selection.startRow, selection.endRow),
    startCol: Math.min(selection.startCol, selection.endCol),
    endCol: Math.max(selection.startCol, selection.endCol),
    mode: selection.mode || "cell"
  };
}

function setHcSelection(startRow, startCol, endRow = startRow, endCol = startCol, mode = "cell") {
  state.hcSelection = { startRow, startCol, endRow, endCol, mode };
}

function isHcCellSelected(rowIndex, colIndex) {
  const selection = normalizeHcSelection();
  return rowIndex >= selection.startRow && rowIndex <= selection.endRow && colIndex >= selection.startCol && colIndex <= selection.endCol;
}

function isHcSelectionTailCell(rowIndex, colIndex) {
  const selection = normalizeHcSelection();
  return rowIndex === selection.endRow && colIndex === selection.endCol;
}

function isSameHcCell(a, b) {
  return a?.rowIndex === b?.rowIndex && a?.colIndex === b?.colIndex;
}

function getActiveHcCell() {
  const selection = normalizeHcSelection();
  return { rowIndex: selection.startRow, colIndex: selection.startCol };
}

function getActiveHcCellMeta() {
  const { rowIndex, colIndex } = getActiveHcCell();
  const month = getHcMonths()[colIndex];
  const code = hcAccountCodes[rowIndex];
  return { rowIndex, colIndex, month, code };
}

function clampHcRowIndex(rowIndex) {
  return Math.min(Math.max(Number(rowIndex) || 0, 0), Math.max(hcAccountCodes.length - 1, 0));
}

function clampHcColIndex(colIndex) {
  return Math.min(Math.max(Number(colIndex) || 0, 0), Math.max(getHcMonths().length - 1, 0));
}

function focusHcCellInput(rowIndex, colIndex, { selectText = true } = {}) {
  const input = els.hcModule.querySelector(`input[data-hc-row="${rowIndex}"][data-hc-col="${colIndex}"]`);
  if (!input) return;
  input.focus();
  if (selectText && typeof input.select === "function") input.select();
  input.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function setActiveHcCell(rowIndex, colIndex, { selectText = true } = {}) {
  const nextRowIndex = clampHcRowIndex(rowIndex);
  const nextColIndex = clampHcColIndex(colIndex);
  setHcSelection(nextRowIndex, nextColIndex);
  state.hcEditCell = { rowIndex: nextRowIndex, colIndex: nextColIndex };
  renderHcModule();
  window.requestAnimationFrame(() => focusHcCellInput(nextRowIndex, nextColIndex, { selectText }));
}

function fillHcSelectionToLastRow() {
  fillHcSelectionRight(getHcMonths().length - 1);
}

function performHcFillToLastMonth(rowIndex, colIndex) {
  setHcSelection(rowIndex, colIndex);
  state.hcEditCell = { rowIndex, colIndex };
  state.hcFillDrag = null;
  fillHcSelectionToLastRow();
  render();
  focusHcCellInput(rowIndex, colIndex, { selectText: false });
}

function getHcSelectionText() {
  const selection = normalizeHcSelection();
  const months = getHcMonths();
  return Array.from({ length: selection.endRow - selection.startRow + 1 }, (_, rowOffset) => {
    const rowIndex = selection.startRow + rowOffset;
    return Array.from({ length: selection.endCol - selection.startCol + 1 }, (_, colOffset) => {
      const colIndex = selection.startCol + colOffset;
      return getHcCellRawValue(hcAccountCodes[rowIndex], months[colIndex]);
    }).join("\t");
  }).join("\n");
}

function adjustFormulaReferences(raw, rowDelta, colDelta) {
  const text = String(raw || "");
  if (!text.startsWith("=") || !colDelta) return text;
  return text.replace(/(A\d+(?:\.\d+)*)@(\d{4}[/-]\d{1,2})/g, (_, code, monthToken) => {
    const normalizedMonth = normalizeHcFormulaMonthToken(monthToken, "");
    if (!normalizedMonth) return `${code}@${monthToken}`;
    return buildHcFormulaReference(code, addMonths(normalizedMonth, colDelta));
  });
}

function fillHcSelectionRight(targetColIndex) {
  const selection = normalizeHcSelection();
  const months = getHcMonths();
  if (targetColIndex <= selection.endCol) return;
  const sourceWidth = selection.endCol - selection.startCol + 1;
  for (let rowIndex = selection.startRow; rowIndex <= selection.endRow; rowIndex += 1) {
    for (let colIndex = selection.endCol + 1; colIndex <= targetColIndex; colIndex += 1) {
      const sourceColIndex = selection.startCol + ((colIndex - selection.endCol - 1) % sourceWidth);
      const targetCode = hcAccountCodes[rowIndex];
      const sourceMonth = months[sourceColIndex];
      const targetMonth = months[colIndex];
      const raw = getHcCellRawValue(targetCode, sourceMonth);
      const adjustedRaw = adjustFormulaReferences(raw, 0, colIndex - sourceColIndex);
      setHcCellRawValue(targetCode, targetMonth, adjustedRaw, { recompute: false });
    }
  }
  recomputeHcForecastValues();
}

function normalizeSheetSelection(selection) {
  return {
    startRow: Math.min(selection.startRow, selection.endRow),
    endRow: Math.max(selection.startRow, selection.endRow),
    startCol: Math.min(selection.startCol, selection.endCol),
    endCol: Math.max(selection.startCol, selection.endCol),
    mode: selection.mode || "cell"
  };
}

function isSheetCellSelected(selection, rowIndex, colIndex) {
  const normalized = normalizeSheetSelection(selection);
  return rowIndex >= normalized.startRow
    && rowIndex <= normalized.endRow
    && colIndex >= normalized.startCol
    && colIndex <= normalized.endCol;
}

function isSheetSelectionTailCell(selection, rowIndex, colIndex) {
  const normalized = normalizeSheetSelection(selection);
  return rowIndex === normalized.endRow && colIndex === normalized.endCol;
}

function getRevenueMonths() {
  return monthRange(state.forecastStart, state.forecastMonths);
}

function getRevenueBlueprint(code) {
  return revenueRowBlueprints[code] || {
    displayName: getTemplateAccountName(code),
    feeType: "收入",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  };
}

function getRevenueReferencePeriodMonths(referencePeriodRaw) {
  const text = String(referencePeriodRaw || "").trim();
  if (text === "近1个月") return 1;
  if (text === "近3个月") return 3;
  if (text === "近6个月") return 6;
  if (text === "近12个月") return 12;
  return 0;
}

function getRevenueRowByCode(code) {
  return state.revenueForecastRows.find((item) => item.code === code) || null;
}

function getRevenueHistoryMonthlyTotals(code) {
  if (code === "A2") {
    return state.flowRows
      .map((row) => [normalizeMonth(row.month), Number(row.revenue || 0)])
      .filter(([month]) => month)
      .sort((a, b) => a[0].localeCompare(b[0]));
  }

  const monthMap = new Map();
  getBoundaryRows()
    .filter((row) => row.code === code)
    .forEach((row) => {
      const month = normalizeMonth(row.month || row.dt);
      if (!month) return;
      monthMap.set(month, (monthMap.get(month) || 0) + Number(row.amount || 0));
    });
  return [...monthMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}

function parseRevenueDriverCode(costDriver) {
  const matched = String(costDriver || "").trim().match(/^(A\d+(?:\.\d+)?)/);
  return matched?.[1] || "";
}

function deriveRevenueSuggestedValueText(code, referencePeriodRaw) {
  const row = getRevenueRowByCode(code) || getRevenueBlueprint(code);
  if (row.computed) return "-";
  const periodMonths = getRevenueReferencePeriodMonths(referencePeriodRaw);
  if (!periodMonths) return "-";
  const monthlyTotals = getRevenueHistoryMonthlyTotals(code);
  if (!monthlyTotals.length) return "-";

  if (String(row.estimateMethodRaw || row.estimateMethod || "").trim() === "费用率") {
    const driverCode = parseRevenueDriverCode(row.costDriver);
    const driverMonthlyTotals = getRevenueHistoryMonthlyTotals(driverCode);
    if (!driverCode || !driverMonthlyTotals.length) return "-";
    const rowSlice = monthlyTotals.slice(-periodMonths);
    const driverMap = new Map(driverMonthlyTotals);
    const numerator = sum(rowSlice.map(([, value]) => Number(value || 0)));
    const denominator = sum(rowSlice.map(([month]) => Number(driverMap.get(month) || 0)));
    return denominator ? ratio(numerator / denominator) : "-";
  }

  const values = monthlyTotals.slice(-periodMonths).map(([, value]) => Number(value || 0));
  return values.length ? integerWithThousands(avg(values)) : "-";
}

function deriveRevenueForecastLogic(methodRaw, fallback = "") {
  const method = String(methodRaw || "").trim();
  if (!method) return fallback;
  if (method === "按实际需求输入或读取excel") return "按实际需求输入或读取excel";
  if (method === "费用率") return "=预估参数*成本动因";
  if (method.startsWith("=")) return "按预估方法公式自动计算";
  return fallback;
}

function getRevenueCellRawValue(code, month) {
  const row = getRevenueRowByCode(code);
  return sanitizeAutoCalculatedRawValue(row, row?.rawValues?.[month] ?? "");
}

function getRevenueCellNumericValue(code, month) {
  const row = getRevenueRowByCode(code);
  return Number(row?.values?.[month] || 0);
}

function getRevenueCellEditValue(code, month) {
  const rawValue = getRevenueCellRawValue(code, month);
  if (rawValue !== "") return rawValue;
  const numericValue = getRevenueCellNumericValue(code, month);
  return numericValue ? compactDecimal(numericValue, 0) : "";
}

function getRevenueCellAddress(rowIndex, colIndex) {
  const code = revenueAccountCodes[rowIndex] || "-";
  const month = formatMonthSlashLabel(getRevenueMonths()[colIndex]) || "-";
  return `${code} · ${month}`;
}

function buildRevenueFormulaReference(code, month) {
  return `${code}@${formatMonthSlashLabel(month)}`;
}

function parseRevenueFormulaReference(reference, fallbackMonth) {
  const [codePart, monthPart] = String(reference || "").split("@");
  return {
    code: String(codePart || "").trim(),
    month: normalizeHcFormulaMonthToken(monthPart, fallbackMonth)
  };
}

function getActiveRevenueFormulaEditor() {
  const activeElement = document.activeElement;
  if (!activeElement) return null;

  if (activeElement.id === "revenue-formula-input") {
    const activeCell = getActiveRevenueCellMeta();
    if (!activeCell.code || !activeCell.month) return null;
    return {
      element: activeElement,
      kind: "formulaBar",
      code: activeCell.code,
      month: activeCell.month,
      rowIndex: activeCell.rowIndex,
      colIndex: activeCell.colIndex
    };
  }

  const cellInput = activeElement.closest?.("input[data-revenue-code][data-revenue-month]");
  if (!cellInput) return null;
  return {
    element: cellInput,
    kind: "cellInput",
    code: cellInput.dataset.revenueCode,
    month: cellInput.dataset.revenueMonth,
    rowIndex: Number(cellInput.dataset.revenueRow),
    colIndex: Number(cellInput.dataset.revenueCol)
  };
}

function syncRevenueFormulaEditorValue(editor, nextValue) {
  updateRevenueValue(editor.code, editor.month, nextValue, { recompute: false });
  const formulaBar = document.getElementById("revenue-formula-input");
  if (formulaBar && editor.kind !== "formulaBar") formulaBar.value = nextValue;
  if (editor.kind === "formulaBar") {
    const sourceInput = els.revenueModule.querySelector(`input[data-revenue-row="${editor.rowIndex}"][data-revenue-col="${editor.colIndex}"]`);
    if (sourceInput) sourceInput.value = nextValue;
  }
}

function tryInsertRevenueFormulaReferenceFromCell(cell) {
  const editor = getActiveRevenueFormulaEditor();
  if (!editor) return false;
  if (!String(editor.element.value || "").trim().startsWith("=")) return false;

  const rowIndex = Number(cell.dataset.revenueRow);
  const colIndex = Number(cell.dataset.revenueCol);
  const code = revenueAccountCodes[rowIndex];
  const month = getRevenueMonths()[colIndex];
  if (!code || !month) return false;

  const reference = buildRevenueFormulaReference(code, month);
  const nextValue = insertTextAtSelection(editor.element, reference);
  syncRevenueFormulaEditorValue(editor, nextValue);
  state.revenueFormulaInsertLock = true;
  return true;
}

function evaluateRevenueFormula(formula, currentRowIndex, currentColIndex, visiting, cache) {
  const month = getRevenueMonths()[currentColIndex];
  const expression = sanitizeFormulaExpression(formula.slice(1))
    .replace(/A\d+(?:\.\d+)*(?:@\d{4}[/-]\d{1,2})?/g, (matchedReference) => {
      const { code, month: refMonth } = parseRevenueFormulaReference(matchedReference, month);
      const refRowIndex = revenueAccountCodes.indexOf(code);
      const refColIndex = getRevenueMonths().indexOf(refMonth);
      if (refRowIndex < 0 || refColIndex < 0 || !refMonth) return "0";
      return String(evaluateRevenueCell(code, refMonth, refRowIndex, refColIndex, visiting, cache));
    });

  if (/[^0-9+\-*/().]/.test(expression)) return 0;
  try {
    const value = Function(`"use strict"; return (${expression || 0});`)();
    return Number.isFinite(Number(value)) ? Number(value) : 0;
  } catch {
    return 0;
  }
}

function resolveRevenueDriverValue(row, month, visiting, cache) {
  const driverCode = parseRevenueDriverCode(row.costDriver);
  if (!driverCode) return 0;
  const refRowIndex = revenueAccountCodes.indexOf(driverCode);
  const refColIndex = getRevenueMonths().indexOf(month);
  if (refRowIndex < 0 || refColIndex < 0) return 0;
  return evaluateRevenueCell(driverCode, month, refRowIndex, refColIndex, visiting, cache);
}

function evaluateRevenueCell(code, month, rowIndex, colIndex, visiting = new Set(), cache = new Map()) {
  const key = `${code}__${month}`;
  if (cache.has(key)) return cache.get(key);
  if (visiting.has(key)) return 0;
  visiting.add(key);

  const row = getRevenueRowByCode(code);
  const rawValue = getRevenueCellRawValue(code, month);
  let value = 0;
  if (!row) {
    value = 0;
  } else if (row.computed) {
    const formula = rawValue || row.estimateMethodRaw;
    value = formula.startsWith("=")
      ? evaluateRevenueFormula(formula, rowIndex, colIndex, visiting, cache)
      : parseNumericInput(formula);
  } else if (isRateEstimateMethod(row.estimateMethodRaw)) {
    value = rawValue.startsWith("=")
      ? evaluateRevenueFormula(rawValue, rowIndex, colIndex, visiting, cache)
      : (resolveRevenueDriverValue(row, month, visiting, cache) * parseRateInput(row.forecastParamRaw));
  } else if (rawValue.startsWith("=")) {
    value = evaluateRevenueFormula(rawValue, rowIndex, colIndex, visiting, cache);
  } else if (rawValue !== "") {
    value = parseNumericInput(rawValue);
  } else {
    value = 0;
  }

  visiting.delete(key);
  cache.set(key, value);
  return value;
}

function recomputeRevenueForecastValues() {
  const months = getRevenueMonths();
  const cache = new Map();
  state.revenueForecastRows.forEach((row, rowIndex) => {
    row.values = row.values || {};
    months.forEach((month, colIndex) => {
      row.values[month] = evaluateRevenueCell(row.code, month, rowIndex, colIndex, new Set(), cache);
    });
  });
}

function setRevenueCellRawValue(code, month, rawValue, { recompute = true } = {}) {
  const row = getRevenueRowByCode(code);
  if (!row || isAutoCalculatedForecastRow(row)) return;
  row.rawValues = row.rawValues || {};
  row.rawValues[month] = normalizeHcRawInput(rawValue);
  if (recompute) recomputeRevenueForecastValues();
}

function normalizeRevenueSelection(selection = state.revenueSelection) {
  return normalizeSheetSelection(selection);
}

function setRevenueSelection(startRow, startCol, endRow = startRow, endCol = startCol, mode = "cell") {
  state.revenueSelection = { startRow, startCol, endRow, endCol, mode };
}

function isRevenueCellSelected(rowIndex, colIndex) {
  return isSheetCellSelected(state.revenueSelection, rowIndex, colIndex);
}

function isRevenueSelectionTailCell(rowIndex, colIndex) {
  return isSheetSelectionTailCell(state.revenueSelection, rowIndex, colIndex);
}

function getActiveRevenueCell() {
  const selection = normalizeRevenueSelection();
  return { rowIndex: selection.startRow, colIndex: selection.startCol };
}

function getActiveRevenueCellMeta() {
  const { rowIndex, colIndex } = getActiveRevenueCell();
  const month = getRevenueMonths()[colIndex];
  const code = revenueAccountCodes[rowIndex];
  return { rowIndex, colIndex, month, code };
}

function clampRevenueRowIndex(rowIndex) {
  return Math.min(Math.max(Number(rowIndex) || 0, 0), Math.max(revenueAccountCodes.length - 1, 0));
}

function clampRevenueColIndex(colIndex) {
  return Math.min(Math.max(Number(colIndex) || 0, 0), Math.max(getRevenueMonths().length - 1, 0));
}

function focusRevenueCellInput(rowIndex, colIndex, { selectText = true } = {}) {
  const input = els.revenueModule.querySelector(`input[data-revenue-row="${rowIndex}"][data-revenue-col="${colIndex}"]`);
  if (!input) return;
  input.focus();
  if (selectText && typeof input.select === "function") input.select();
  input.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function setActiveRevenueCell(rowIndex, colIndex, { selectText = true } = {}) {
  const nextRowIndex = clampRevenueRowIndex(rowIndex);
  const nextColIndex = clampRevenueColIndex(colIndex);
  setRevenueSelection(nextRowIndex, nextColIndex);
  state.revenueEditCell = { rowIndex: nextRowIndex, colIndex: nextColIndex };
  renderRevenueModule();
  window.requestAnimationFrame(() => focusRevenueCellInput(nextRowIndex, nextColIndex, { selectText }));
}

function getRevenueSelectionText() {
  const selection = normalizeRevenueSelection();
  const months = getRevenueMonths();
  return Array.from({ length: selection.endRow - selection.startRow + 1 }, (_, rowOffset) => {
    const rowIndex = selection.startRow + rowOffset;
    return Array.from({ length: selection.endCol - selection.startCol + 1 }, (_, colOffset) => {
      const colIndex = selection.startCol + colOffset;
      return getRevenueCellRawValue(revenueAccountCodes[rowIndex], months[colIndex]);
    }).join("\t");
  }).join("\n");
}

function fillRevenueSelectionRight(targetColIndex) {
  const selection = normalizeRevenueSelection();
  const months = getRevenueMonths();
  if (targetColIndex <= selection.endCol) return;
  const sourceWidth = selection.endCol - selection.startCol + 1;
  for (let rowIndex = selection.startRow; rowIndex <= selection.endRow; rowIndex += 1) {
    const code = revenueAccountCodes[rowIndex];
    if (getRevenueRowByCode(code)?.computed) continue;
    for (let colIndex = selection.endCol + 1; colIndex <= targetColIndex; colIndex += 1) {
      const sourceColIndex = selection.startCol + ((colIndex - selection.endCol - 1) % sourceWidth);
      const sourceMonth = months[sourceColIndex];
      const targetMonth = months[colIndex];
      const rawValue = getRevenueCellRawValue(code, sourceMonth);
      const adjustedRawValue = adjustFormulaReferences(rawValue, 0, colIndex - sourceColIndex);
      setRevenueCellRawValue(code, targetMonth, adjustedRawValue, { recompute: false });
    }
  }
  recomputeRevenueForecastValues();
}

function fillRevenueSelectionToLastRow() {
  fillRevenueSelectionRight(getRevenueMonths().length - 1);
}

function getRuleEditableColumns() {
  return [
    { field: "expectedValue", label: "预期值" },
    { field: "growthRate", label: "增长设置" }
  ];
}

function getRenderableRules() {
  const scope = getActiveScope();
  const rules = getRulesByScope(scope).filter((rule) => !isManualModuleAccountCode(rule.code));
  const search = state.filters.search.trim().toLowerCase();
  return rules.filter((rule) => {
    const sourceOk = state.filters.source === "all" || rule.source === state.filters.source;
    const typeOk = state.filters.type === "all" || rule.type === state.filters.type;
    const searchOk = !search || `${rule.code} ${rule.account}`.toLowerCase().includes(search);
    return sourceOk && typeOk && searchOk;
  });
}

function getRuleCellRawValue(rule, field) {
  const rawField = `${field}Raw`;
  return rule?.[rawField] ?? String(rule?.[field] ?? "");
}

function getRuleCellNumericValue(rule, field) {
  return Number(rule?.[field] || 0);
}

function getRuleCellAddress(rowIndex, colIndex) {
  return `${colIndexToLetters(colIndex)}${rowIndex + 1}`;
}

function evaluateRuleFormula(formula, rows, visiting, cache) {
  const columns = getRuleEditableColumns();
  const expression = sanitizeFormulaExpression(formula.slice(1)).replace(/([$]?)([A-Z]+)([$]?)(\d+)/g, (_, absCol, letters, absRow, rowNumberText) => {
    const refColIndex = lettersToColIndex(letters);
    const refRowIndex = Number(rowNumberText) - 1;
    if (refColIndex < 0 || refColIndex >= columns.length || refRowIndex < 0 || refRowIndex >= rows.length) {
      return "0";
    }
    return String(evaluateRuleCell(rows[refRowIndex], columns[refColIndex].field, refRowIndex, refColIndex, rows, visiting, cache));
  });

  if (/[^0-9+\-*/().]/.test(expression)) return 0;
  try {
    const value = Function(`"use strict"; return (${expression || 0});`)();
    return Number.isFinite(Number(value)) ? Number(value) : 0;
  } catch {
    return 0;
  }
}

function evaluateRuleCell(rule, field, rowIndex, colIndex, rows, visiting = new Set(), cache = new Map()) {
  const key = `${rule?.key || rowIndex}__${field}`;
  if (cache.has(key)) return cache.get(key);
  if (visiting.has(key)) return 0;
  visiting.add(key);
  const rawValue = getRuleCellRawValue(rule, field);
  const value = rawValue.startsWith("=")
    ? evaluateRuleFormula(rawValue, rows, visiting, cache)
    : parseNumericInput(rawValue);
  visiting.delete(key);
  cache.set(key, value);
  return value;
}

function recomputeRuleGridValues(rows = getRenderableRules()) {
  const columns = getRuleEditableColumns();
  const cache = new Map();
  rows.forEach((rule, rowIndex) => {
    columns.forEach((column, colIndex) => {
      rule[column.field] = evaluateRuleCell(rule, column.field, rowIndex, colIndex, rows, new Set(), cache);
    });
  });
  applyLinkedRuleSummaries(getActiveScope());
}

function setRuleCellRawValue(rule, field, rawValue, { recompute = true, rows = getRenderableRules() } = {}) {
  if (!rule) return;
  rule[`${field}Raw`] = normalizeHcRawInput(rawValue);
  if (recompute) recomputeRuleGridValues(rows);
}

function normalizeRulesSelection(selection = state.rulesSelection) {
  return normalizeSheetSelection(selection);
}

function setRulesSelection(startRow, startCol, endRow = startRow, endCol = startCol, mode = "cell") {
  state.rulesSelection = { startRow, startCol, endRow, endCol, mode };
}

function isRuleCellSelected(rowIndex, colIndex) {
  return isSheetCellSelected(state.rulesSelection, rowIndex, colIndex);
}

function isRulesSelectionTailCell(rowIndex, colIndex) {
  return isSheetSelectionTailCell(state.rulesSelection, rowIndex, colIndex);
}

function getActiveRuleCell() {
  const selection = normalizeRulesSelection();
  return { rowIndex: selection.startRow, colIndex: selection.startCol };
}

function getActiveRuleCellMeta() {
  const { rowIndex, colIndex } = getActiveRuleCell();
  const rows = getRenderableRules();
  const columns = getRuleEditableColumns();
  return {
    rowIndex,
    colIndex,
    rule: rows[rowIndex] || null,
    field: columns[colIndex]?.field || "expectedValue"
  };
}

function clampRuleRowIndex(rowIndex) {
  return Math.min(Math.max(Number(rowIndex) || 0, 0), Math.max(getRenderableRules().length - 1, 0));
}

function clampRuleColIndex(colIndex) {
  return Math.min(Math.max(Number(colIndex) || 0, 0), Math.max(getRuleEditableColumns().length - 1, 0));
}

function focusRuleCellInput(rowIndex, colIndex, { selectText = true } = {}) {
  const input = els.rulesBody.querySelector(`input[data-rule-row="${rowIndex}"][data-rule-col="${colIndex}"]`);
  if (!input) return;
  input.focus();
  if (selectText && typeof input.select === "function") input.select();
  input.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function setActiveRuleCell(rowIndex, colIndex, { selectText = true } = {}) {
  const nextRowIndex = clampRuleRowIndex(rowIndex);
  const nextColIndex = clampRuleColIndex(colIndex);
  setRulesSelection(nextRowIndex, nextColIndex);
  state.rulesEditCell = { rowIndex: nextRowIndex, colIndex: nextColIndex };
  renderRules();
  window.requestAnimationFrame(() => focusRuleCellInput(nextRowIndex, nextColIndex, { selectText }));
}

function getRuleSelectionText(rows = getRenderableRules()) {
  const selection = normalizeRulesSelection();
  const columns = getRuleEditableColumns();
  return Array.from({ length: selection.endRow - selection.startRow + 1 }, (_, rowOffset) => {
    const rowIndex = selection.startRow + rowOffset;
    const rule = rows[rowIndex];
    return Array.from({ length: selection.endCol - selection.startCol + 1 }, (_, colOffset) => {
      const colIndex = selection.startCol + colOffset;
      return getRuleCellRawValue(rule, columns[colIndex].field);
    }).join("\t");
  }).join("\n");
}

function fillRulesSelectionDown(targetRowIndex) {
  const rows = getRenderableRules();
  const selection = normalizeRulesSelection();
  const columns = getRuleEditableColumns();
  if (targetRowIndex <= selection.endRow) return;
  const sourceHeight = selection.endRow - selection.startRow + 1;
  for (let rowIndex = selection.endRow + 1; rowIndex <= targetRowIndex; rowIndex += 1) {
    for (let colIndex = selection.startCol; colIndex <= selection.endCol; colIndex += 1) {
      const sourceRowIndex = selection.startRow + ((rowIndex - selection.endRow - 1) % sourceHeight);
      const targetRule = rows[rowIndex];
      const sourceRule = rows[sourceRowIndex];
      if (!targetRule || !sourceRule) continue;
      const field = columns[colIndex].field;
      const rawValue = getRuleCellRawValue(sourceRule, field);
      const adjustedRawValue = adjustFormulaReferences(rawValue, rowIndex - sourceRowIndex, 0);
      setRuleCellRawValue(targetRule, field, adjustedRawValue, { recompute: false, rows });
    }
  }
  recomputeRuleGridValues(rows);
}

function fillRulesSelectionToLastRow() {
  fillRulesSelectionDown(getRenderableRules().length - 1);
}

function isHcAccountCode(code) {
  return hcAccountCodes.includes(String(code || "").trim());
}

function isRevenueAccountCode(code) {
  return revenueAccountCodes.includes(String(code || "").trim());
}

function isCostAccountCode(code) {
  return costAccountCodes.includes(String(code || "").trim());
}

function isMarketingAccountCode(code) {
  return marketingAccountCodes.includes(String(code || "").trim());
}

function isManualModuleAccountCode(code) {
  return isHcAccountCode(code) || isRevenueAccountCode(code) || isCostAccountCode(code) || isMarketingAccountCode(code);
}

function getTemplateAccountName(code) {
  const normalizedCode = String(code || "").trim();
  return loadedTemplateLayoutNameLookup.get(normalizedCode)
    || templateAccounts.find((item) => item.code === normalizedCode)?.name
    || "";
}

function getBoundaryRows() {
  if (state.importMode === "post") return state.actualRows.post;
  return [...state.actualRows.pre, ...state.actualRows.allocated];
}

function getHcSourceRows() {
  if (state.importMode === "post") return state.actualRows.post;
  return state.actualRows.pre.length ? state.actualRows.pre : state.actualRows.post;
}

function getRowsMaxMonth(rows) {
  const months = rows.map((row) => normalizeMonth(row.month || row.dt)).filter(Boolean).sort();
  return months.at(-1) || "";
}

function getNextForecastStart(rows) {
  const maxMonth = getRowsMaxMonth(rows);
  return maxMonth ? addMonths(maxMonth, 1) : state.forecastStart;
}

function getLatestHc(rows) {
  const maxMonth = getRowsMaxMonth(rows);
  if (!maxMonth) return 1;
  const hcs = rows
    .filter((row) => normalizeMonth(row.month || row.dt) === maxMonth)
    .map((row) => Number(row.hc || 0))
    .filter((value) => value > 0);
  return hcs[0] || 1;
}

function syncRuleBoundaries() {
  Object.values(state.rules).forEach((rules) => {
    rules.forEach((rule) => {
      rule.startMonth = state.forecastStart;
      rule.growthStartMonth = addMonths(state.forecastStart, 6);
      if (rule.endMonth !== "tail" && rule.endMonth < state.forecastStart) {
        rule.endMonth = "tail";
      }
    });
  });
}

function rebuildHcForecastRows() {
  const months = getHcMonths();
  const sourceRows = getHcSourceRows();
  const previousRows = new Map(
    state.hcForecastRows.map((row) => [row.code, row])
  );

  state.hcForecastRows = hcAccountCodes.map((code) => {
    const blueprint = getHcBlueprint(code);
    const previousRow = previousRows.get(code) || {};
    const values = {};
    const rawValues = {};
    const codeRows = sourceRows.filter((row) => row.code === code);
    const latestMonth = getRowsMaxMonth(codeRows);
    const defaultValue = latestMonth
      ? sum(codeRows.filter((row) => normalizeMonth(row.month || row.dt) === latestMonth).map((row) => Number(row.amount || 0)))
      : 0;
    const historySummary = getHcHistorySummary(code);
    const defaultRawValue = blueprint.computed
      ? blueprint.estimateMethod
      : String(defaultValue || "");

    months.forEach((month) => {
      const previousRawValue = previousRow.rawValues?.[month];
      rawValues[month] = previousRawValue ?? defaultRawValue;
      values[month] = previousRow.values?.[month] ?? defaultValue;
    });

    return {
      code,
      account: blueprint.displayName || getTemplateAccountName(code),
      feeType: previousRow.feeType || blueprint.feeType,
      estimateMethodRaw: previousRow.estimateMethodRaw || blueprint.estimateMethod,
      costDriver: previousRow.costDriver || blueprint.costDriver,
      suggestedLogic: previousRow.suggestedLogic || blueprint.suggestedLogic,
      referencePeriodRaw: previousRow.referencePeriodRaw || blueprint.referencePeriod || "",
      suggestedValueText: previousRow.referencePeriodRaw || blueprint.referencePeriod
        ? deriveHcSuggestedValueText(code, previousRow.referencePeriodRaw || blueprint.referencePeriod || "")
        : "-",
      forecastParamRaw: previousRow.forecastParamRaw || blueprint.forecastParam || "",
      forecastLogicRaw: previousRow.forecastLogicRaw || deriveHcForecastLogic(previousRow.estimateMethodRaw || blueprint.estimateMethod, blueprint.forecastLogic || ""),
      computed: Boolean(blueprint.computed),
      rawValues,
      values
    };
  });
  recomputeHcForecastValues();
}

function rebuildRevenueForecastParts() {
  rebuildRevenueForecastRows();
}

function rebuildRevenueForecastRows({ resetRevenueFromFlow = false } = {}) {
  const months = getRevenueMonths();
  const previousRows = new Map(
    state.revenueForecastRows.map((row) => [row.code, row])
  );
  const flowValueMap = new Map(
    state.flowRows.map((row) => [normalizeMonth(row.month), Number(row.revenue || 0)])
  );

  state.revenueForecastRows = revenueAccountCodes.map((code) => {
    const blueprint = getRevenueBlueprint(code);
    const previousRow = previousRows.get(code) || {};
    const rawValues = {};
    const values = {};
    const referencePeriodRaw = previousRow.referencePeriodRaw || blueprint.referencePeriod || "";
    const estimateMethodRaw = previousRow.estimateMethodRaw || blueprint.estimateMethod || "";

    months.forEach((month) => {
      const defaultRawValue = code === "A2"
        ? (flowValueMap.get(month) ? String(flowValueMap.get(month)) : "")
        : (blueprint.computed || isRateEstimateMethod(estimateMethodRaw) ? String(blueprint.computed ? (estimateMethodRaw || blueprint.estimateMethod || "") : "") : "");
      const previousRawValue = previousRow.rawValues?.[month];
      rawValues[month] = blueprint.computed
        ? String(estimateMethodRaw || blueprint.estimateMethod || "")
        : (resetRevenueFromFlow && code === "A2"
          ? defaultRawValue
          : (isRateEstimateMethod(estimateMethodRaw)
            ? sanitizeAutoCalculatedRawValue({ estimateMethodRaw }, previousRawValue)
            : (previousRawValue ?? defaultRawValue)));
      values[month] = previousRow.values?.[month] ?? 0;
    });

    const nextRow = {
      code,
      account: blueprint.displayName || getTemplateAccountName(code),
      feeType: previousRow.feeType || blueprint.feeType,
      estimateMethodRaw,
      costDriver: previousRow.costDriver || blueprint.costDriver || "",
      suggestedLogic: previousRow.suggestedLogic || blueprint.suggestedLogic || "",
      referencePeriodRaw,
      suggestedValueText: "-",
      forecastParamRaw: previousRow.forecastParamRaw || blueprint.forecastParam || "",
      forecastLogicRaw: previousRow.forecastLogicRaw || deriveRevenueForecastLogic(estimateMethodRaw, blueprint.forecastLogic || ""),
      computed: Boolean(blueprint.computed),
      rawValues,
      values
    };
    nextRow.suggestedValueText = deriveRevenueSuggestedValueText(code, nextRow.referencePeriodRaw);
    if (!nextRow.forecastParamRaw && nextRow.estimateMethodRaw === "费用率" && nextRow.suggestedValueText !== "-") {
      nextRow.forecastParamRaw = nextRow.suggestedValueText;
    }
    if (nextRow.computed) {
      months.forEach((month) => {
        nextRow.rawValues[month] = nextRow.estimateMethodRaw;
      });
    }
    return nextRow;
  });
  recomputeRevenueForecastValues();
}

function collectFlowLabelCandidates(row) {
  return Object.entries(row || {})
    .filter(([, value]) => typeof value === "string" && String(value).trim())
    .map(([key, value]) => ({
      key: String(key || "").trim().toLowerCase(),
      value: String(value || "").trim()
    }));
}

function normalizeFlowRegionLabel(row) {
  const candidates = collectFlowLabelCandidates(row);
  for (const candidate of candidates) {
    const value = candidate.value.toLowerCase();
    if (/cn|china|中国|大陆|国服/.test(value)) return "CN";
    if (/海外|oversea|overseas|intl|international|global|hk|tw|jp|kr|sea|na|eu|港台|港澳台|日韩|欧美/.test(value)) return "海外";
  }
  return "";
}

function normalizeFlowPlatformLabel(row) {
  const candidates = collectFlowLabelCandidates(row);
  for (const candidate of candidates) {
    const value = candidate.value.toLowerCase();
    if (/console|ps|switch|xbox|主机/.test(value)) return "Console";
    if (/pc|steam|epic|wegame|客户端|端游/.test(value)) return "PC";
    if (/mobile|ios|android|google play|app store|移动|手游/.test(value)) return "移动";
  }
  return "";
}

function normalizeFlowRows(rows) {
  const monthTotals = new Map();
  rows.forEach((row) => {
    const month = normalizeMonth(row.month || row.dt);
    if (!month) return;
    const revenue = Number(row.revenue || row.amount || row.grossRevenue || 0);
    const existing = monthTotals.get(month) || {
      month,
      revenue: 0,
      regionBreakdown: {},
      platformBreakdown: {}
    };
    existing.revenue += revenue;
    const regionLabel = normalizeFlowRegionLabel(row);
    const platformLabel = normalizeFlowPlatformLabel(row);
    if (regionLabel) {
      existing.regionBreakdown[regionLabel] = Number(existing.regionBreakdown[regionLabel] || 0) + revenue;
    }
    if (platformLabel) {
      existing.platformBreakdown[platformLabel] = Number(existing.platformBreakdown[platformLabel] || 0) + revenue;
    }
    monthTotals.set(month, existing);
  });

  return [...monthTotals.values()].sort((a, b) => a.month.localeCompare(b.month));
}

function applyImportedFlowRows(rows, fileName, meta = {}) {
  const normalizedRows = normalizeFlowRows(rows);
  state.flowRows = normalizedRows;
  state.flowImportName = fileName;
  state.flowImportRowCount = normalizedRows.length;
  state.flowImportMatchedChannels = Array.isArray(meta.matchedChannels) ? meta.matchedChannels : [];
  state.flowImportUnmatchedChannels = Array.isArray(meta.unmatchedChannels) ? meta.unmatchedChannels : [];
  state.flowImportSourceFormat = meta.sourceFormat || "";
  rebuildRevenueForecastRows({ resetRevenueFromFlow: true });
}

function getRevenueForecastSummary(months = monthRange(state.forecastStart, state.forecastMonths)) {
  return months.map((month) => {
    const grossRevenue = getRevenueCellNumericValue("A2", month);
    const vatAmount = getRevenueCellNumericValue("A3", month);
    const platformAmount = getRevenueCellNumericValue("A4", month);
    const netRevenue = getRevenueCellNumericValue("A5", month);
    const vatRatio = grossRevenue ? vatAmount / grossRevenue : 0;
    const platformRatio = grossRevenue ? platformAmount / grossRevenue : 0;

    return {
      month,
      grossRevenue,
      vatAmount,
      platformAmount,
      netRevenue,
      vatRatio,
      platformRatio,
      netRatio: grossRevenue ? netRevenue / grossRevenue : 0
    };
  });
}

function getCostMonths() {
  return monthRange(state.forecastStart, state.forecastMonths);
}

function getCostBlueprint(code) {
  return costRowBlueprints[code] || {
    displayName: getTemplateAccountName(code),
    feeType: "变动费用",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近1个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  };
}

function getCostReferencePeriodMonths(referencePeriodRaw) {
  const text = String(referencePeriodRaw || "").trim();
  if (text === "近1个月") return 1;
  if (text === "近3个月") return 3;
  if (text === "近6个月") return 6;
  if (text === "近12个月") return 12;
  return 0;
}

function getCostRowByCode(code) {
  return state.costForecastRows.find((item) => item.code === code) || null;
}

function getHistoricalRevenueByMonth() {
  const monthMap = new Map();
  getBoundaryRows().forEach((row) => {
    const month = normalizeMonth(row.month || row.dt);
    const revenue = Number(row.revenue || 0);
    if (!month || !revenue) return;
    if (!monthMap.has(month)) monthMap.set(month, revenue);
  });
  return [...monthMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}

function getCostHistoryMonthlyTotals(code) {
  if (code === "A2") return getHistoricalRevenueByMonth();
  const monthMap = new Map();
  getBoundaryRows()
    .filter((row) => row.code === code)
    .forEach((row) => {
      const month = normalizeMonth(row.month || row.dt);
      if (!month) return;
      monthMap.set(month, (monthMap.get(month) || 0) + Number(row.amount || 0));
    });
  return [...monthMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}

function deriveCostSuggestedValueText(code, referencePeriodRaw) {
  const row = getCostRowByCode(code) || getCostBlueprint(code);
  if (row.computed) return "-";
  const periodMonths = getCostReferencePeriodMonths(referencePeriodRaw);
  if (!periodMonths) return "-";
  const monthlyTotals = getCostHistoryMonthlyTotals(code);
  if (!monthlyTotals.length) return "-";

  if (String(row.estimateMethodRaw || row.estimateMethod || "").trim() === "费用率") {
    const driverMonthlyTotals = getCostHistoryMonthlyTotals("A2");
    const rowSlice = monthlyTotals.slice(-periodMonths);
    const driverMap = new Map(driverMonthlyTotals);
    const numerator = sum(rowSlice.map(([, value]) => Number(value || 0)));
    const denominator = sum(rowSlice.map(([month]) => Number(driverMap.get(month) || 0)));
    return denominator ? ratio(numerator / denominator) : "-";
  }

  const values = monthlyTotals.slice(-periodMonths).map(([, value]) => Number(value || 0));
  return values.length ? integerWithThousands(avg(values)) : "-";
}

function deriveCostForecastLogic(methodRaw, fallback = "") {
  const method = String(methodRaw || "").trim();
  if (!method) return fallback;
  if (method === "按实际需求输入或读取excel") return "按实际需求输入或读取excel";
  if (method === "费用率") return "=预估参数*成本动因";
  if (method.startsWith("=")) return "按预估方法公式自动计算";
  return fallback;
}

function getCostCellRawValue(code, month) {
  const row = getCostRowByCode(code);
  return sanitizeAutoCalculatedRawValue(row, row?.rawValues?.[month] ?? "");
}

function getCostCellNumericValue(code, month) {
  if (code === "A2") return getRevenueCellNumericValue("A2", month);
  if (code === "A5") return getRevenueCellNumericValue("A5", month);
  const row = getCostRowByCode(code);
  return Number(row?.values?.[month] || 0);
}

function getCostCellEditValue(code, month) {
  const rawValue = getCostCellRawValue(code, month);
  if (rawValue !== "") return rawValue;
  const numericValue = getCostCellNumericValue(code, month);
  return numericValue ? compactDecimal(numericValue, 0) : "";
}

function getCostCellAddress(rowIndex, colIndex) {
  const code = costAccountCodes[rowIndex] || "-";
  const month = formatMonthSlashLabel(getCostMonths()[colIndex]) || "-";
  return `${code} · ${month}`;
}

function buildCostFormulaReference(code, month) {
  return `${code}@${formatMonthSlashLabel(month)}`;
}

function parseCostFormulaReference(reference, fallbackMonth) {
  const [codePart, monthPart] = String(reference || "").split("@");
  return {
    code: String(codePart || "").trim(),
    month: normalizeHcFormulaMonthToken(monthPart, fallbackMonth)
  };
}

function getActiveCostFormulaEditor() {
  const activeElement = document.activeElement;
  if (!activeElement) return null;
  if (activeElement.id === "cost-formula-input") {
    const activeCell = getActiveCostCellMeta();
    if (!activeCell.code || !activeCell.month) return null;
    return { element: activeElement, kind: "formulaBar", ...activeCell };
  }
  const cellInput = activeElement.closest?.("input[data-cost-code][data-cost-month]");
  if (!cellInput) return null;
  return {
    element: cellInput,
    kind: "cellInput",
    code: cellInput.dataset.costCode,
    month: cellInput.dataset.costMonth,
    rowIndex: Number(cellInput.dataset.costRow),
    colIndex: Number(cellInput.dataset.costCol)
  };
}

function syncCostFormulaEditorValue(editor, nextValue) {
  updateCostValue(editor.code, editor.month, nextValue, { recompute: false });
  const formulaBar = document.getElementById("cost-formula-input");
  if (formulaBar && editor.kind !== "formulaBar") formulaBar.value = nextValue;
  if (editor.kind === "formulaBar") {
    const sourceInput = els.costModule.querySelector(`input[data-cost-row="${editor.rowIndex}"][data-cost-col="${editor.colIndex}"]`);
    if (sourceInput) sourceInput.value = nextValue;
  }
}

function tryInsertCostFormulaReferenceFromCell(cell) {
  const editor = getActiveCostFormulaEditor();
  if (!editor) return false;
  if (!String(editor.element.value || "").trim().startsWith("=")) return false;
  const rowIndex = Number(cell.dataset.costRow);
  const colIndex = Number(cell.dataset.costCol);
  const code = costAccountCodes[rowIndex];
  const month = getCostMonths()[colIndex];
  if (!code || !month) return false;
  const reference = buildCostFormulaReference(code, month);
  const nextValue = insertTextAtSelection(editor.element, reference);
  syncCostFormulaEditorValue(editor, nextValue);
  state.costFormulaInsertLock = true;
  return true;
}

function evaluateCostReference(code, month, visiting, cache) {
  if (code === "A2") return getRevenueCellNumericValue("A2", month);
  if (code === "A5") return getRevenueCellNumericValue("A5", month);
  const refRowIndex = costAccountCodes.indexOf(code);
  const refColIndex = getCostMonths().indexOf(month);
  if (refRowIndex < 0 || refColIndex < 0) return 0;
  return evaluateCostCell(code, month, refRowIndex, refColIndex, visiting, cache);
}

function evaluateCostFormula(formula, currentRowIndex, currentColIndex, visiting, cache) {
  const month = getCostMonths()[currentColIndex];
  const expression = sanitizeFormulaExpression(formula.slice(1))
    .replace(/A\d+(?:\.\d+)*(?:@\d{4}[/-]\d{1,2})?/g, (matchedReference) => {
      const { code, month: refMonth } = parseCostFormulaReference(matchedReference, month);
      if (!refMonth) return "0";
      return String(evaluateCostReference(code, refMonth, visiting, cache));
    });
  if (/[^0-9+\-*/().]/.test(expression)) return 0;
  try {
    const value = Function(`"use strict"; return (${expression || 0});`)();
    return Number.isFinite(Number(value)) ? Number(value) : 0;
  } catch {
    return 0;
  }
}

function resolveCostDriverValue(row, month) {
  if (parseRevenueDriverCode(row.costDriver) === "A2") return getRevenueCellNumericValue("A2", month);
  return 0;
}

function evaluateCostCell(code, month, rowIndex, colIndex, visiting = new Set(), cache = new Map()) {
  const key = `${code}__${month}`;
  if (cache.has(key)) return cache.get(key);
  if (visiting.has(key)) return 0;
  visiting.add(key);
  const row = getCostRowByCode(code);
  const rawValue = getCostCellRawValue(code, month);
  let value = 0;
  if (!row) {
    value = 0;
  } else if (row.computed) {
    const formula = rawValue || row.estimateMethodRaw;
    value = formula.startsWith("=") ? evaluateCostFormula(formula, rowIndex, colIndex, visiting, cache) : parseNumericInput(formula);
  } else if (isRateEstimateMethod(row.estimateMethodRaw)) {
    value = rawValue.startsWith("=")
      ? evaluateCostFormula(rawValue, rowIndex, colIndex, visiting, cache)
      : (resolveCostDriverValue(row, month) * parseRateInput(row.forecastParamRaw));
  } else if (rawValue.startsWith("=")) {
    value = evaluateCostFormula(rawValue, rowIndex, colIndex, visiting, cache);
  } else if (rawValue !== "") {
    value = parseNumericInput(rawValue);
  } else {
    value = 0;
  }
  visiting.delete(key);
  cache.set(key, value);
  return value;
}

function recomputeCostForecastValues() {
  const months = getCostMonths();
  const cache = new Map();
  state.costForecastRows.forEach((row, rowIndex) => {
    row.values = row.values || {};
    months.forEach((month, colIndex) => {
      row.values[month] = evaluateCostCell(row.code, month, rowIndex, colIndex, new Set(), cache);
    });
  });
}

function setCostCellRawValue(code, month, rawValue, { recompute = true } = {}) {
  const row = getCostRowByCode(code);
  if (!row || isAutoCalculatedForecastRow(row)) return;
  row.rawValues = row.rawValues || {};
  row.rawValues[month] = normalizeHcRawInput(rawValue);
  if (recompute) recomputeCostForecastValues();
}

function normalizeCostSelection(selection = state.costSelection) {
  return normalizeSheetSelection(selection);
}

function setCostSelection(startRow, startCol, endRow = startRow, endCol = startCol, mode = "cell") {
  state.costSelection = { startRow, startCol, endRow, endCol, mode };
}

function isCostCellSelected(rowIndex, colIndex) {
  return isSheetCellSelected(state.costSelection, rowIndex, colIndex);
}

function isCostSelectionTailCell(rowIndex, colIndex) {
  return isSheetSelectionTailCell(state.costSelection, rowIndex, colIndex);
}

function getActiveCostCell() {
  const selection = normalizeCostSelection();
  return { rowIndex: selection.startRow, colIndex: selection.startCol };
}

function getActiveCostCellMeta() {
  const { rowIndex, colIndex } = getActiveCostCell();
  const month = getCostMonths()[colIndex];
  const code = costAccountCodes[rowIndex];
  return { rowIndex, colIndex, month, code };
}

function clampCostRowIndex(rowIndex) {
  return Math.min(Math.max(Number(rowIndex) || 0, 0), Math.max(costAccountCodes.length - 1, 0));
}

function clampCostColIndex(colIndex) {
  return Math.min(Math.max(Number(colIndex) || 0, 0), Math.max(getCostMonths().length - 1, 0));
}

function focusCostCellInput(rowIndex, colIndex, { selectText = true } = {}) {
  const input = els.costModule.querySelector(`input[data-cost-row="${rowIndex}"][data-cost-col="${colIndex}"]`);
  if (!input) return;
  input.focus();
  if (selectText && typeof input.select === "function") input.select();
  input.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function setActiveCostCell(rowIndex, colIndex, { selectText = true } = {}) {
  const nextRowIndex = clampCostRowIndex(rowIndex);
  const nextColIndex = clampCostColIndex(colIndex);
  setCostSelection(nextRowIndex, nextColIndex);
  state.costEditCell = { rowIndex: nextRowIndex, colIndex: nextColIndex };
  renderCostModule();
  window.requestAnimationFrame(() => focusCostCellInput(nextRowIndex, nextColIndex, { selectText }));
}

function getCostSelectionText() {
  const selection = normalizeCostSelection();
  const months = getCostMonths();
  return Array.from({ length: selection.endRow - selection.startRow + 1 }, (_, rowOffset) => {
    const rowIndex = selection.startRow + rowOffset;
    return Array.from({ length: selection.endCol - selection.startCol + 1 }, (_, colOffset) => {
      const colIndex = selection.startCol + colOffset;
      return getCostCellRawValue(costAccountCodes[rowIndex], months[colIndex]);
    }).join("\t");
  }).join("\n");
}

function fillCostSelectionRight(targetColIndex) {
  const selection = normalizeCostSelection();
  const months = getCostMonths();
  if (targetColIndex <= selection.endCol) return;
  const sourceWidth = selection.endCol - selection.startCol + 1;
  for (let rowIndex = selection.startRow; rowIndex <= selection.endRow; rowIndex += 1) {
    const code = costAccountCodes[rowIndex];
    if (getCostRowByCode(code)?.computed) continue;
    for (let colIndex = selection.endCol + 1; colIndex <= targetColIndex; colIndex += 1) {
      const sourceColIndex = selection.startCol + ((colIndex - selection.endCol - 1) % sourceWidth);
      const sourceMonth = months[sourceColIndex];
      const targetMonth = months[colIndex];
      const rawValue = getCostCellRawValue(code, sourceMonth);
      const adjustedRawValue = adjustFormulaReferences(rawValue, 0, colIndex - sourceColIndex);
      setCostCellRawValue(code, targetMonth, adjustedRawValue, { recompute: false });
    }
  }
  recomputeCostForecastValues();
}

function fillCostSelectionToLastRow() {
  fillCostSelectionRight(getCostMonths().length - 1);
}

function rebuildCostForecastRows() {
  const months = getCostMonths();
  const previousRows = new Map(state.costForecastRows.map((row) => [row.code, row]));
  state.costForecastRows = costAccountCodes.map((code) => {
    const blueprint = getCostBlueprint(code);
    const previousRow = previousRows.get(code) || {};
    const rawValues = {};
    const values = {};
    const historyTotals = getCostHistoryMonthlyTotals(code);
    const latestValue = historyTotals.at(-1)?.[1] || 0;
    const estimateMethodRaw = previousRow.estimateMethodRaw || blueprint.estimateMethod || "";
    const referencePeriodRaw = previousRow.referencePeriodRaw || blueprint.referencePeriod || "";
    months.forEach((month) => {
      const defaultRawValue = blueprint.computed
        ? String(estimateMethodRaw || blueprint.estimateMethod || "")
        : (isRateEstimateMethod(estimateMethodRaw) ? "" : (latestValue ? String(latestValue) : ""));
      rawValues[month] = blueprint.computed
        ? String(estimateMethodRaw || blueprint.estimateMethod || "")
        : (isRateEstimateMethod(estimateMethodRaw)
          ? sanitizeAutoCalculatedRawValue({ estimateMethodRaw }, previousRow.rawValues?.[month])
          : (previousRow.rawValues?.[month] ?? defaultRawValue));
      values[month] = previousRow.values?.[month] ?? 0;
    });
    const nextRow = {
      code,
      account: blueprint.displayName || getTemplateAccountName(code),
      feeType: previousRow.feeType || blueprint.feeType,
      estimateMethodRaw,
      costDriver: previousRow.costDriver || blueprint.costDriver || "",
      suggestedLogic: previousRow.suggestedLogic || blueprint.suggestedLogic || "",
      referencePeriodRaw,
      suggestedValueText: "-",
      forecastParamRaw: previousRow.forecastParamRaw || blueprint.forecastParam || "",
      forecastLogicRaw: previousRow.forecastLogicRaw || deriveCostForecastLogic(estimateMethodRaw, blueprint.forecastLogic || ""),
      computed: Boolean(blueprint.computed),
      rawValues,
      values
    };
    nextRow.suggestedValueText = deriveCostSuggestedValueText(code, nextRow.referencePeriodRaw);
    if (!nextRow.forecastParamRaw && nextRow.estimateMethodRaw === "费用率" && nextRow.suggestedValueText !== "-") {
      nextRow.forecastParamRaw = nextRow.suggestedValueText;
    }
    if (nextRow.computed) {
      months.forEach((month) => { nextRow.rawValues[month] = nextRow.estimateMethodRaw; });
    }
    return nextRow;
  });
  recomputeCostForecastValues();
}

function getMarketingMonths() {
  return monthRange(state.forecastStart, state.forecastMonths);
}

function getMarketingBlueprint(code) {
  return marketingRowBlueprints[code] || {
    displayName: getTemplateAccountName(code),
    feeType: "变动费用-市场类",
    estimateMethod: "按实际需求输入或读取excel",
    costDriver: "按需",
    suggestedLogic: "=参考期间均值",
    referencePeriod: "近3个月",
    forecastParam: "",
    forecastLogic: "按实际需求输入或读取excel",
    computed: false
  };
}

function getMarketingReferencePeriodMonths(referencePeriodRaw) {
  const text = String(referencePeriodRaw || "").trim();
  if (text === "近1个月") return 1;
  if (text === "近3个月") return 3;
  if (text === "近6个月") return 6;
  if (text === "近12个月") return 12;
  return 0;
}

function getMarketingRowByCode(code) {
  return state.marketingForecastRows.find((item) => item.code === code) || null;
}

function getMarketingHistoryMonthlyTotals(code) {
  if (code === "A2") return getHistoricalRevenueByMonth();
  const monthMap = new Map();
  getBoundaryRows()
    .filter((row) => row.code === code)
    .forEach((row) => {
      const month = normalizeMonth(row.month || row.dt);
      if (!month) return;
      monthMap.set(month, (monthMap.get(month) || 0) + Number(row.amount || 0));
    });
  return [...monthMap.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}

function resolveMarketingHistoryDriverCode(costDriver) {
  if (String(costDriver || "").includes("A11")) return "A11";
  if (String(costDriver || "").includes("A2")) return "A2";
  return "";
}

function deriveMarketingSuggestedValueText(code, referencePeriodRaw) {
  const row = getMarketingRowByCode(code) || getMarketingBlueprint(code);
  if (row.computed) return "-";
  const periodMonths = getMarketingReferencePeriodMonths(referencePeriodRaw);
  if (!periodMonths) return "-";
  const monthlyTotals = getMarketingHistoryMonthlyTotals(code);
  if (!monthlyTotals.length) return "-";

  if (String(row.estimateMethodRaw || row.estimateMethod || "").trim() === "费用率") {
    const driverCode = resolveMarketingHistoryDriverCode(row.costDriver);
    const driverMonthlyTotals = getMarketingHistoryMonthlyTotals(driverCode);
    const rowSlice = monthlyTotals.slice(-periodMonths);
    const driverMap = new Map(driverMonthlyTotals);
    const numerator = sum(rowSlice.map(([, value]) => Number(value || 0)));
    const denominator = sum(rowSlice.map(([month]) => Number(driverMap.get(month) || 0)));
    return denominator ? ratio(numerator / denominator) : "-";
  }

  const values = monthlyTotals.slice(-periodMonths).map(([, value]) => Number(value || 0));
  return values.length ? integerWithThousands(avg(values)) : "-";
}

function deriveMarketingForecastLogic(methodRaw, fallback = "") {
  const method = String(methodRaw || "").trim();
  if (!method) return fallback;
  if (method === "按实际需求输入或读取excel") return "按实际需求输入或读取excel";
  if (method === "费用率") return "=预估参数*成本动因";
  if (method.startsWith("=")) return "按预估方法公式自动计算";
  return fallback;
}

function isRateEstimateMethod(methodRaw) {
  return String(methodRaw || "").trim() === "费用率";
}

function isAutoCalculatedForecastRow(row) {
  return Boolean(row?.computed) || isRateEstimateMethod(row?.estimateMethodRaw);
}

function sanitizeAutoCalculatedRawValue(row, rawValue) {
  const text = String(rawValue ?? "");
  if (!isRateEstimateMethod(row?.estimateMethodRaw)) return text;
  return text.trim().startsWith("=") ? text : "";
}

function shouldLockForecastParam(row) {
  return Boolean(row?.computed) || String(row?.forecastLogicRaw || "").trim() === "按实际需求输入或读取excel";
}

function shouldUseNativeCopy(event) {
  const target = event.target;
  const targetInput = target?.closest?.("input, textarea");
  if (targetInput && typeof targetInput.selectionStart === "number" && typeof targetInput.selectionEnd === "number" && targetInput.selectionStart !== targetInput.selectionEnd) {
    return true;
  }
  const activeElement = document.activeElement;
  if (activeElement?.matches?.("input, textarea")
    && typeof activeElement.selectionStart === "number"
    && typeof activeElement.selectionEnd === "number"
    && activeElement.selectionStart !== activeElement.selectionEnd) {
    return true;
  }
  const selectedText = window.getSelection?.()?.toString?.() || "";
  return Boolean(selectedText.trim());
}

function refocusMetaFieldInput(container, selector) {
  window.requestAnimationFrame(() => {
    const input = container?.querySelector(selector);
    if (!input || typeof input.focus !== "function") return;
    input.focus();
    const caret = String(input.value || "").length;
    if (typeof input.setSelectionRange === "function") {
      input.setSelectionRange(caret, caret);
    }
  });
}

function getMarketingCellRawValue(code, month) {
  const row = getMarketingRowByCode(code);
  return sanitizeAutoCalculatedRawValue(row, row?.rawValues?.[month] ?? "");
}

function getMarketingCellNumericValue(code, month) {
  if (code === "A2") return getRevenueCellNumericValue("A2", month);
  if (code === "A9") return getCostCellNumericValue("A9", month);
  const row = getMarketingRowByCode(code);
  return Number(row?.values?.[month] || 0);
}

function getMarketingCellEditValue(code, month) {
  const rawValue = getMarketingCellRawValue(code, month);
  if (rawValue !== "") return rawValue;
  const numericValue = getMarketingCellNumericValue(code, month);
  return numericValue ? compactDecimal(numericValue, 0) : "";
}

function getMarketingCellAddress(rowIndex, colIndex) {
  const code = marketingAccountCodes[rowIndex] || "-";
  const month = formatMonthSlashLabel(getMarketingMonths()[colIndex]) || "-";
  return `${code} · ${month}`;
}

function buildMarketingFormulaReference(code, month) {
  return `${code}@${formatMonthSlashLabel(month)}`;
}

function parseMarketingFormulaReference(reference, fallbackMonth) {
  const [codePart, monthPart] = String(reference || "").split("@");
  return {
    code: String(codePart || "").trim(),
    month: normalizeHcFormulaMonthToken(monthPart, fallbackMonth)
  };
}

function getActiveMarketingFormulaEditor() {
  const activeElement = document.activeElement;
  if (!activeElement) return null;
  if (activeElement.id === "marketing-formula-input") {
    const activeCell = getActiveMarketingCellMeta();
    if (!activeCell.code || !activeCell.month) return null;
    return { element: activeElement, kind: "formulaBar", ...activeCell };
  }
  const cellInput = activeElement.closest?.("input[data-marketing-code][data-marketing-month]");
  if (!cellInput) return null;
  return {
    element: cellInput,
    kind: "cellInput",
    code: cellInput.dataset.marketingCode,
    month: cellInput.dataset.marketingMonth,
    rowIndex: Number(cellInput.dataset.marketingRow),
    colIndex: Number(cellInput.dataset.marketingCol)
  };
}

function syncMarketingFormulaEditorValue(editor, nextValue) {
  updateMarketingValue(editor.code, editor.month, nextValue, { recompute: false });
  const formulaBar = document.getElementById("marketing-formula-input");
  if (formulaBar && editor.kind !== "formulaBar") formulaBar.value = nextValue;
  if (editor.kind === "formulaBar") {
    const sourceInput = els.marketingModule.querySelector(`input[data-marketing-row="${editor.rowIndex}"][data-marketing-col="${editor.colIndex}"]`);
    if (sourceInput) sourceInput.value = nextValue;
  }
}

function tryInsertMarketingFormulaReferenceFromCell(cell) {
  const editor = getActiveMarketingFormulaEditor();
  if (!editor) return false;
  if (!String(editor.element.value || "").trim().startsWith("=")) return false;
  const rowIndex = Number(cell.dataset.marketingRow);
  const colIndex = Number(cell.dataset.marketingCol);
  const code = marketingAccountCodes[rowIndex];
  const month = getMarketingMonths()[colIndex];
  if (!code || !month) return false;
  const reference = buildMarketingFormulaReference(code, month);
  const nextValue = insertTextAtSelection(editor.element, reference);
  syncMarketingFormulaEditorValue(editor, nextValue);
  state.marketingFormulaInsertLock = true;
  return true;
}

function evaluateMarketingReference(code, month, visiting, cache) {
  if (code === "A2") return getRevenueCellNumericValue("A2", month);
  if (code === "A9") return getCostCellNumericValue("A9", month);
  const refRowIndex = marketingAccountCodes.indexOf(code);
  const refColIndex = getMarketingMonths().indexOf(month);
  if (refRowIndex < 0 || refColIndex < 0) return 0;
  return evaluateMarketingCell(code, month, refRowIndex, refColIndex, visiting, cache);
}

function evaluateMarketingFormula(formula, currentRowIndex, currentColIndex, visiting, cache) {
  const month = getMarketingMonths()[currentColIndex];
  const expression = sanitizeFormulaExpression(formula.slice(1))
    .replace(/A\d+(?:\.\d+)*(?:@\d{4}[/-]\d{1,2})?/g, (matchedReference) => {
      const { code, month: refMonth } = parseMarketingFormulaReference(matchedReference, month);
      if (!refMonth) return "0";
      return String(evaluateMarketingReference(code, refMonth, visiting, cache));
    });
  if (/[^0-9+\-*/().]/.test(expression)) return 0;
  try {
    const value = Function(`"use strict"; return (${expression || 0});`)();
    return Number.isFinite(Number(value)) ? Number(value) : 0;
  } catch {
    return 0;
  }
}

function resolveMarketingDriverValue(row, month, visiting, cache) {
  const driverCode = resolveMarketingHistoryDriverCode(row.costDriver);
  if (!driverCode) return 0;
  if (driverCode === "A2") return getRevenueCellNumericValue("A2", month);
  const refRowIndex = marketingAccountCodes.indexOf(driverCode);
  const refColIndex = getMarketingMonths().indexOf(month);
  if (refRowIndex < 0 || refColIndex < 0) return 0;
  return evaluateMarketingCell(driverCode, month, refRowIndex, refColIndex, visiting, cache);
}

function evaluateMarketingCell(code, month, rowIndex, colIndex, visiting = new Set(), cache = new Map()) {
  const key = `${code}__${month}`;
  if (cache.has(key)) return cache.get(key);
  if (visiting.has(key)) return 0;
  visiting.add(key);
  const row = getMarketingRowByCode(code);
  const rawValue = getMarketingCellRawValue(code, month);
  let value = 0;
  if (!row) {
    value = 0;
  } else if (row.computed) {
    const formula = rawValue || row.estimateMethodRaw;
    value = formula.startsWith("=") ? evaluateMarketingFormula(formula, rowIndex, colIndex, visiting, cache) : parseNumericInput(formula);
  } else if (isRateEstimateMethod(row.estimateMethodRaw)) {
    value = rawValue.startsWith("=")
      ? evaluateMarketingFormula(rawValue, rowIndex, colIndex, visiting, cache)
      : (resolveMarketingDriverValue(row, month, visiting, cache) * parseRateInput(row.forecastParamRaw));
  } else if (rawValue.startsWith("=")) {
    value = evaluateMarketingFormula(rawValue, rowIndex, colIndex, visiting, cache);
  } else if (rawValue !== "") {
    value = parseNumericInput(rawValue);
  } else {
    value = 0;
  }
  visiting.delete(key);
  cache.set(key, value);
  return value;
}

function recomputeMarketingForecastValues() {
  const months = getMarketingMonths();
  const cache = new Map();
  state.marketingForecastRows.forEach((row, rowIndex) => {
    row.values = row.values || {};
    months.forEach((month, colIndex) => {
      row.values[month] = evaluateMarketingCell(row.code, month, rowIndex, colIndex, new Set(), cache);
    });
  });
}

function setMarketingCellRawValue(code, month, rawValue, { recompute = true } = {}) {
  const row = getMarketingRowByCode(code);
  if (!row || isAutoCalculatedForecastRow(row)) return;
  row.rawValues = row.rawValues || {};
  row.rawValues[month] = normalizeHcRawInput(rawValue);
  if (recompute) recomputeMarketingForecastValues();
}

function normalizeMarketingSelection(selection = state.marketingSelection) {
  return normalizeSheetSelection(selection);
}

function setMarketingSelection(startRow, startCol, endRow = startRow, endCol = startCol, mode = "cell") {
  state.marketingSelection = { startRow, startCol, endRow, endCol, mode };
}

function isMarketingCellSelected(rowIndex, colIndex) {
  return isSheetCellSelected(state.marketingSelection, rowIndex, colIndex);
}

function isMarketingSelectionTailCell(rowIndex, colIndex) {
  return isSheetSelectionTailCell(state.marketingSelection, rowIndex, colIndex);
}

function getActiveMarketingCell() {
  const selection = normalizeMarketingSelection();
  return { rowIndex: selection.startRow, colIndex: selection.startCol };
}

function getActiveMarketingCellMeta() {
  const { rowIndex, colIndex } = getActiveMarketingCell();
  const month = getMarketingMonths()[colIndex];
  const code = marketingAccountCodes[rowIndex];
  return { rowIndex, colIndex, month, code };
}

function clampMarketingRowIndex(rowIndex) {
  return Math.min(Math.max(Number(rowIndex) || 0, 0), Math.max(marketingAccountCodes.length - 1, 0));
}

function clampMarketingColIndex(colIndex) {
  return Math.min(Math.max(Number(colIndex) || 0, 0), Math.max(getMarketingMonths().length - 1, 0));
}

function focusMarketingCellInput(rowIndex, colIndex, { selectText = true } = {}) {
  const input = els.marketingModule.querySelector(`input[data-marketing-row="${rowIndex}"][data-marketing-col="${colIndex}"]`);
  if (!input) return;
  input.focus();
  if (selectText && typeof input.select === "function") input.select();
  input.scrollIntoView({ block: "nearest", inline: "nearest" });
}

function setActiveMarketingCell(rowIndex, colIndex, { selectText = true } = {}) {
  const nextRowIndex = clampMarketingRowIndex(rowIndex);
  const nextColIndex = clampMarketingColIndex(colIndex);
  setMarketingSelection(nextRowIndex, nextColIndex);
  state.marketingEditCell = { rowIndex: nextRowIndex, colIndex: nextColIndex };
  renderMarketingModule();
  window.requestAnimationFrame(() => focusMarketingCellInput(nextRowIndex, nextColIndex, { selectText }));
}

function getMarketingSelectionText() {
  const selection = normalizeMarketingSelection();
  const months = getMarketingMonths();
  return Array.from({ length: selection.endRow - selection.startRow + 1 }, (_, rowOffset) => {
    const rowIndex = selection.startRow + rowOffset;
    return Array.from({ length: selection.endCol - selection.startCol + 1 }, (_, colOffset) => {
      const colIndex = selection.startCol + colOffset;
      return getMarketingCellRawValue(marketingAccountCodes[rowIndex], months[colIndex]);
    }).join("\t");
  }).join("\n");
}

function fillMarketingSelectionRight(targetColIndex) {
  const selection = normalizeMarketingSelection();
  const months = getMarketingMonths();
  if (targetColIndex <= selection.endCol) return;
  const sourceWidth = selection.endCol - selection.startCol + 1;
  for (let rowIndex = selection.startRow; rowIndex <= selection.endRow; rowIndex += 1) {
    const code = marketingAccountCodes[rowIndex];
    if (getMarketingRowByCode(code)?.computed) continue;
    for (let colIndex = selection.endCol + 1; colIndex <= targetColIndex; colIndex += 1) {
      const sourceColIndex = selection.startCol + ((colIndex - selection.endCol - 1) % sourceWidth);
      const sourceMonth = months[sourceColIndex];
      const targetMonth = months[colIndex];
      const rawValue = getMarketingCellRawValue(code, sourceMonth);
      const adjustedRawValue = adjustFormulaReferences(rawValue, 0, colIndex - sourceColIndex);
      setMarketingCellRawValue(code, targetMonth, adjustedRawValue, { recompute: false });
    }
  }
  recomputeMarketingForecastValues();
}

function fillMarketingSelectionToLastRow() {
  fillMarketingSelectionRight(getMarketingMonths().length - 1);
}

function rebuildMarketingForecastRows() {
  const months = getMarketingMonths();
  const previousRows = new Map(state.marketingForecastRows.map((row) => [row.code, row]));
  state.marketingForecastRows = marketingAccountCodes.map((code) => {
    const blueprint = getMarketingBlueprint(code);
    const previousRow = previousRows.get(code) || {};
    const rawValues = {};
    const values = {};
    const historyTotals = getMarketingHistoryMonthlyTotals(code);
    const latestValue = historyTotals.at(-1)?.[1] || 0;
    const estimateMethodRaw = previousRow.estimateMethodRaw || blueprint.estimateMethod || "";
    const referencePeriodRaw = previousRow.referencePeriodRaw || blueprint.referencePeriod || "";
    months.forEach((month) => {
      const defaultRawValue = blueprint.computed
        ? String(estimateMethodRaw || blueprint.estimateMethod || "")
        : (isRateEstimateMethod(estimateMethodRaw) ? "" : (latestValue ? String(latestValue) : ""));
      rawValues[month] = blueprint.computed
        ? String(estimateMethodRaw || blueprint.estimateMethod || "")
        : (isRateEstimateMethod(estimateMethodRaw)
          ? sanitizeAutoCalculatedRawValue({ estimateMethodRaw }, previousRow.rawValues?.[month])
          : (previousRow.rawValues?.[month] ?? defaultRawValue));
      values[month] = previousRow.values?.[month] ?? 0;
    });
    const nextRow = {
      code,
      account: blueprint.displayName || getTemplateAccountName(code),
      feeType: previousRow.feeType || blueprint.feeType,
      estimateMethodRaw,
      costDriver: previousRow.costDriver || blueprint.costDriver || "",
      suggestedLogic: previousRow.suggestedLogic || blueprint.suggestedLogic || "",
      referencePeriodRaw,
      suggestedValueText: "-",
      forecastParamRaw: previousRow.forecastParamRaw || blueprint.forecastParam || "",
      forecastLogicRaw: previousRow.forecastLogicRaw || deriveMarketingForecastLogic(estimateMethodRaw, blueprint.forecastLogic || ""),
      computed: Boolean(blueprint.computed),
      rawValues,
      values
    };
    nextRow.suggestedValueText = deriveMarketingSuggestedValueText(code, nextRow.referencePeriodRaw);
    if (!nextRow.forecastParamRaw && nextRow.estimateMethodRaw === "费用率" && nextRow.suggestedValueText !== "-") {
      nextRow.forecastParamRaw = nextRow.suggestedValueText;
    }
    if (nextRow.computed) months.forEach((month) => { nextRow.rawValues[month] = nextRow.estimateMethodRaw; });
    return nextRow;
  });
  recomputeMarketingForecastValues();
}

function createRevenueSummaryBucket(label) {
  return {
    label,
    months: [],
    grossRevenue: 0,
    vatAmount: 0,
    platformAmount: 0,
    netRevenue: 0
  };
}

function appendRevenueSummaryBucket(bucket, row) {
  bucket.months.push(row.month);
  bucket.grossRevenue += Number(row.grossRevenue || 0);
  bucket.vatAmount += Number(row.vatAmount || 0);
  bucket.platformAmount += Number(row.platformAmount || 0);
  bucket.netRevenue += Number(row.netRevenue || 0);
}

function buildRevenueNaturalYearSummaries(summaryRows) {
  const yearMap = new Map();
  summaryRows.forEach((row) => {
    const year = String(row.month || "").slice(0, 4) || "-";
    if (!yearMap.has(year)) yearMap.set(year, createRevenueSummaryBucket(year));
    appendRevenueSummaryBucket(yearMap.get(year), row);
  });
  return [...yearMap.values()];
}

function buildRevenueGlYearSummaries(summaryRows, goLiveDate) {
  const goLive = parseInputDate(goLiveDate);
  if (!goLive) return [];

  const goLiveMonth = new Date(goLive.getFullYear(), goLive.getMonth(), 1);
  const buckets = Array.from({ length: 5 }, (_, index) => createRevenueSummaryBucket(`GL${index + 1}`));

  summaryRows.forEach((row) => {
    const monthDate = monthKeyToDate(row.month);
    const diff = monthDiff(goLiveMonth, monthDate);
    if (!Number.isFinite(diff) || diff < 0 || diff >= 60) return;
    const bucketIndex = Math.floor(diff / 12);
    appendRevenueSummaryBucket(buckets[bucketIndex], row);
  });

  return buckets.filter((bucket) => bucket.months.length);
}

function renderRevenueAnnualSummaryTableRows(rows, emptyText) {
  if (!rows.length) {
    return `<tr><td colspan="6" class="table-note">${emptyText}</td></tr>`;
  }
  return rows.map((row) => `
    <tr>
      <td>${row.label}</td>
      <td>${formatCoverageMonths(row.months)}</td>
      <td>${compactMoney(row.grossRevenue)}</td>
      <td>${compactMoney(row.vatAmount)}</td>
      <td>${compactMoney(row.platformAmount)}</td>
      <td>${compactMoney(row.netRevenue)}</td>
    </tr>
  `).join("");
}

function formatMoneyWithRatio(amount, share) {
  return `
    <div class="value-with-ratio">
      <strong>${compactMoney(amount)}</strong>
      <span>占流水 ${ratio(share)}</span>
    </div>
  `;
}

function applyBoundaryDefaults() {
  const rows = getBoundaryRows();
  state.forecastStart = getNextForecastStart(rows);
  state.hcStart = getLatestHc(rows);
  if (els.forecastStart) {
    els.forecastStart.value = state.forecastStart;
  }
  rebuildHcForecastRows();
  rebuildRevenueForecastParts();
  rebuildCostForecastRows();
  rebuildMarketingForecastRows();
  syncRuleBoundaries();
}

function normalizePostRows(rows) {
  return rows
    .map((row) => ({
      display_name: row.display_name || row.displayName || row.project_name || "",
      dt: normalizeMonth(row.dt || row.month),
      fcst_acc_code: row.fcst_acc_code || row.code || "",
      fcst_acc_name: row.fcst_acc_name || row.account || "",
      update_dt: String(row.update_dt || row.updateDt || ""),
      amount: Number(row.amount || 0),
      hc: Number(row.hc || 0),
      revenue: Number(row.revenue || 0)
    }))
    .filter((row) => row.dt && row.fcst_acc_code && row.fcst_acc_name);
}

function buildPostPivot(rows, fallbackDisplayName) {
  const normalizedRows = normalizePostRows(rows);
  const dtColumns = [...new Set(normalizedRows.map((row) => row.dt))].sort();
  const rowMap = new Map();

  normalizedRows.forEach((row) => {
    const key = row.fcst_acc_code;
    if (!key) return;
    if (!rowMap.has(key)) {
      rowMap.set(key, { code: row.fcst_acc_code, name: row.fcst_acc_name, indent: 0, values: Object.fromEntries(dtColumns.map((dt) => [dt, null])) });
    }
    const matched = rowMap.get(key);
    matched.values[row.dt] = (matched.values[row.dt] ?? 0) + row.amount;
  });

  const templateCodes = new Set(state.templateLayoutRows.map((item) => item.code));
  const rowsByTemplate = state.templateLayoutRows.map((tpl) => {
    const matched = rowMap.get(tpl.code);
    return {
      code: tpl.code,
      name: tpl.name || getTemplateAccountName(tpl.code) || matched?.name || "",
      indent: tpl.indent || 0,
      values: matched?.values || Object.fromEntries(dtColumns.map((dt) => [dt, null]))
    };
  });

  const unmatchedRows = [...rowMap.values()]
    .filter((row) => !templateCodes.has(row.code))
    .sort((a, b) => `${a.code}__${a.name}`.localeCompare(`${b.code}__${b.name}`, "zh-CN"));

  const displayName = normalizedRows.find((row) => row.display_name)?.display_name || fallbackDisplayName || "未命名项目";
  const maxDt = dtColumns.at(-1) || "-";

  return { displayName, maxDt, dtColumns, rows: [...rowsByTemplate, ...unmatchedRows] };
}

function formatPivotName(row) {
  const { code, name, indent = 0 } = row;
  if (/^\s/.test(name)) return name;
  if (indent > 0) return `${"\u3000".repeat(indent)}${name}`;
  const depth = Math.max((String(code || "").match(/\./g) || []).length, 0);
  return `${"\u3000".repeat(depth)}${name}`;
}

function formatPivotValue(value) {
  if (value === null || value === undefined) return "";
  return new Intl.NumberFormat("zh-CN", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  }).format(value);
}

function inferDriver(account) {
  if (/薪酬|工资|社保|福利|salary|bonus|performance/i.test(account)) return "HC";
  if (/投放|服务器|带宽|分摊|分成|渠道|支付/.test(account)) return "流水";
  return "HC";
}

function inferSource(account) {
  return /分摊|内包|集团管理服务费/.test(account) ? "分摊值" : "项目组直属";
}

function inferType(account, monthlyAmounts) {
  if (/薪酬|工资|房租|折旧|办公|salary|bonus|performance/i.test(account)) return "fixed";
  const mean = avg(monthlyAmounts);
  const variance = avg(monthlyAmounts.map((item) => (item - mean) ** 2));
  const volatility = mean ? Math.sqrt(variance) / mean : 0;
  return volatility > 0.18 ? "variable" : "fixed";
}

function getOrderedRuleTemplateCodes() {
  return state.templateLayoutRows
    .map((row) => row.code)
    .filter((code, index, list) => !isManualModuleAccountCode(code) && list.indexOf(code) === index);
}

function createRuleFromGroup(group, forcedSource, fallbackCode, fallbackAccount) {
  const sample = group[0] || {};
  const monthlyAmounts = group.map((item) => Number(item.amount || 0));
  const account = sample.account || fallbackAccount || "";
  const code = sample.code || fallbackCode || "";
  const source = sample.source || forcedSource || inferSource(account);
  const driver = inferDriver(account);
  const type = inferType(account, monthlyAmounts);
  const method = inferMethod(type, driver);
  const suggestedValue = monthlyAmounts.length
    ? (type === "fixed"
      ? avg(monthlyAmounts)
      : avg(group.map((item) => item.amount / Math.max(driver === "HC" ? item.hc || 1 : item.revenue || 1, 1))))
    : 0;

  return {
    key: `${source}__${code || ""}`,
    code,
    account,
    source,
    type,
    driver,
    method,
    suggestedValue,
    expectedValue: suggestedValue,
    expectedValueRaw: String(suggestedValue),
    startMonth: state.forecastStart,
    endMonth: "tail",
    growthRate: 0.03,
    growthRateRaw: String(0.03),
    growthStartMonth: addMonths(state.forecastStart, 6)
  };
}

function getRuleModuleByCode(code) {
  return ruleModuleCodeMap.get(String(code || "").trim()) || "admin-other-estimate";
}

function getRuleModuleDefinition(moduleKey) {
  return ruleModuleDefinitions.find((module) => module.key === moduleKey) || ruleModuleDefinitions.at(-1);
}

function isLinkedRuleCode(code) {
  return linkedRuleCodes.has(String(code || "").trim());
}

function isRuleEditable(rule, field) {
  if (!rule) return false;
  if (!isLinkedRuleCode(rule.code)) return true;
  return false;
}

function getLinkedRuleLabel(code) {
  return linkedRuleDefinitions[String(code || "").trim()]?.label || "联动汇总";
}

function getRuleDefaultPreset(code) {
  return ruleDefaultPresets[String(code || "").trim()] || null;
}

function applyRulePreset(rule, { preserveExpectedValue = false } = {}) {
  if (!rule) return rule;
  const preset = getRuleDefaultPreset(rule.code);
  if (!preset) return rule;
  rule.type = preset.type;
  rule.driver = preset.driver;
  rule.method = preset.method;
  rule.growthRate = preset.growthRate;
  rule.growthRateRaw = String(preset.growthRate);
  if (!preserveExpectedValue && preset.method === "rate" && rule.suggestedValue && !String(rule.expectedValueRaw || "").trim()) {
    rule.expectedValue = rule.suggestedValue;
    rule.expectedValueRaw = String(rule.suggestedValue);
  }
  return rule;
}

function getRenderableRuleEntries() {
  const filtered = getRenderableRules();
  return filtered.map((rule, rowIndex) => ({
    rowIndex,
    rule,
    moduleKey: getRuleModuleByCode(rule.code)
  }));
}

function normalizeTemplateLayoutRows(rows) {
  const sourceRows = Array.isArray(rows) ? rows : [];
  const sourceMap = new Map(sourceRows.map((row) => [row.code, row]));
  const mergedRows = defaultTemplateLayoutRows.map((row) => ({
    code: row.code,
    name: sourceMap.get(row.code)?.name || getTemplateAccountName(row.code),
    enName: sourceMap.get(row.code)?.enName || "",
    indent: sourceMap.get(row.code)?.indent ?? row.indent
  }));
  const knownCodes = new Set(mergedRows.map((row) => row.code));
  const extraRows = sourceRows
    .filter((row) => row.code && !knownCodes.has(row.code))
    .map((row) => ({
      code: row.code,
      name: row.name || getTemplateAccountName(row.code),
      enName: row.enName || "",
      indent: row.indent ?? 0
    }));
  return [...mergedRows, ...extraRows];
}

function inferMethod(type, driver) {
  if (type === "fixed") return "fixed";
  return driver === "HC" ? "unit" : "rate";
}

function buildRules(rows, forcedSource) {
  const groups = new Map();
  rows.forEach((row) => {
    const rowSource = forcedSource || row.source;
    const code = String(row.code || "").trim();
    const account = row.account || getTemplateAccountName(code);
    if (!code) return;
    const key = `${rowSource}__${code || ""}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ ...row, code, account, source: rowSource });
  });
  const orderedRules = [];
  const consumedKeys = new Set();

  getOrderedRuleTemplateCodes().forEach((code) => {
    const account = getTemplateAccountName(code);
    const source = forcedSource || inferSource(account);
    const key = `${source}__${code || ""}`;
    if (groups.has(key)) {
      orderedRules.push(createRuleFromGroup(groups.get(key), forcedSource, code, account));
      consumedKeys.add(key);
      return;
    }
    const matchedEntry = [...groups.entries()].find(([groupKey]) => groupKey.includes(`__${code}__`));
    if (matchedEntry) {
      orderedRules.push(createRuleFromGroup(matchedEntry[1], forcedSource, code, account));
      consumedKeys.add(matchedEntry[0]);
      return;
    }
    orderedRules.push(createRuleFromGroup([], forcedSource, code, account));
  });

  const extraRules = [...groups.entries()]
    .filter(([key]) => !consumedKeys.has(key))
    .map(([, group]) => createRuleFromGroup(group, forcedSource))
    .sort((a, b) => `${a.code}__${a.account}`.localeCompare(`${b.code}__${b.account}`, "zh-CN"));

  return [...orderedRules, ...extraRules];
}

function computeRuleMonthlyAmount(rule, month, monthIndex, months, hcMap, grossRevenueMap) {
  if (!rule) return 0;
  const active = month >= rule.startMonth && (rule.endMonth === "tail" || month <= rule.endMonth);
  if (!active) return 0;
  let amount = Number(rule.expectedValue || 0);
  if (rule.type === "fixed") {
    if (rule.growthStartMonth && month >= rule.growthStartMonth) {
      const growthStartIndex = months.indexOf(rule.growthStartMonth);
      const growthIndex = Math.max(monthIndex - growthStartIndex + 1, 1);
      amount = amount * (1 + Number(rule.growthRate || 0)) ** growthIndex;
    }
    return amount;
  }
  const driverValue = rule.driver === "HC" ? Number(hcMap[month] || 0) : Number(grossRevenueMap[month] || 0);
  return driverValue * amount;
}

function getCostForecastMaps(months = getCostMonths()) {
  return {
    otherRevenueMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A97", month)])),
    insourcingRevenueMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A91", month)])),
    revenueShareMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A6", month)])),
    serverCostMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A7", month)])),
    salesTaxMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A8", month)])),
    royaltyMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A54", month)])),
    grossProfitMap: Object.fromEntries(months.map((month) => [month, getCostCellNumericValue("A9", month)]))
  };
}

function getMarketingForecastMaps(months = getMarketingMonths()) {
  return {
    salesMarketingMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A10", month)])),
    uaMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A11", month)])),
    uaMaterialMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A11.1", month)])),
    uaTestMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A11.2", month)])),
    brandingMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A12", month)])),
    userOperationMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A12.1", month)])),
    publishingProfitMap: Object.fromEntries(months.map((month) => [month, getMarketingCellNumericValue("A99", month)]))
  };
}

function buildMonthlyProfitBridge(months, rules, hcMap, grossRevenueMap, netRevenueMap, costMaps, marketingMaps) {
  const leafRules = rules.filter((rule) => !isLinkedRuleCode(rule.code));
  const monthlyLeafValues = Object.fromEntries(months.map((month) => [month, {}]));

  leafRules.forEach((rule) => {
    months.forEach((month, monthIndex) => {
      const amount = computeRuleMonthlyAmount(rule, month, monthIndex, months, hcMap, grossRevenueMap);
      monthlyLeafValues[month][rule.code] = (monthlyLeafValues[month][rule.code] || 0) + amount;
    });
  });

  const adminOtherCodes = [...new Set(
    leafRules
      .filter((rule) => getRuleModuleByCode(rule.code) === "admin-other-estimate" && rule.code !== "A13")
      .map((rule) => rule.code)
  )];

  const getValue = (month, code, visiting = new Set()) => {
    const normalizedCode = String(code || "").trim();
    if (normalizedCode === "A5") return Number(netRevenueMap[month] || 0);
    if (normalizedCode === "A97") return Number(costMaps.otherRevenueMap[month] || 0);
    if (normalizedCode === "A91") return Number(costMaps.insourcingRevenueMap[month] || 0);
    if (normalizedCode === "A6") return Number(costMaps.revenueShareMap[month] || 0);
    if (normalizedCode === "A7") return Number(costMaps.serverCostMap[month] || 0);
    if (normalizedCode === "A8") return Number(costMaps.salesTaxMap[month] || 0);
    if (normalizedCode === "A54") return Number(costMaps.royaltyMap[month] || 0);
    if (normalizedCode === "A9") return Number(costMaps.grossProfitMap[month] || 0);
    if (normalizedCode === "A10") return Number(marketingMaps.salesMarketingMap[month] || 0);
    if (normalizedCode === "A99") return Number(marketingMaps.publishingProfitMap[month] || 0);
    if (!isLinkedRuleCode(normalizedCode)) return Number(monthlyLeafValues[month][normalizedCode] || 0);
    if (visiting.has(normalizedCode)) return 0;
    visiting.add(normalizedCode);
    const definition = linkedRuleDefinitions[normalizedCode];
    const ctx = {
      netRevenue: Number(netRevenueMap[month] || 0),
      adminOtherCodes,
      get: (depCode) => getValue(month, depCode, new Set(visiting)),
      sum: (codes) => codes.reduce((total, depCode) => total + getValue(month, depCode, new Set(visiting)), 0)
    };
    const value = Number(definition?.compute(ctx) || 0);
    visiting.delete(normalizedCode);
    return value;
  };

  return months.map((month) => ({
    month,
    a5: Number(netRevenueMap[month] || 0),
    a97: getValue(month, "A97"),
    a91: getValue(month, "A91"),
    a9: getValue(month, "A9"),
    a10: getValue(month, "A10"),
    a99: getValue(month, "A99"),
    a14: getValue(month, "A14"),
    a18: getValue(month, "A18"),
    a13: getValue(month, "A13")
  }));
}

function applyLinkedRuleSummaries(scope) {
  const rules = getRulesByScope(scope).filter((rule) => !isManualModuleAccountCode(rule.code));
  if (!rules.length) return;

  const months = monthRange(state.forecastStart, state.forecastMonths);
  const revenueSummary = getRevenueForecastSummary(months);
  const grossRevenueMap = Object.fromEntries(revenueSummary.map((row) => [row.month, row.grossRevenue]));
  const netRevenueMap = Object.fromEntries(revenueSummary.map((row) => [row.month, row.netRevenue]));
  const costMaps = getCostForecastMaps(months);
  const marketingMaps = getMarketingForecastMaps(months);
  const hcTotalRow = state.hcForecastRows.find((row) => row.code === "A1");
  const hcMap = Object.fromEntries(months.map((month) => [month, Number(hcTotalRow?.values?.[month] || 0)]));
  const adminOtherCodes = [...new Set(
    rules
      .filter((rule) => getRuleModuleByCode(rule.code) === "admin-other-estimate" && rule.code !== "A13" && !isLinkedRuleCode(rule.code))
      .map((rule) => rule.code)
  )];

  const baseValues = new Map();
  rules.forEach((rule) => {
    if (isLinkedRuleCode(rule.code)) return;
    const normalizedValue = rule.type === "fixed"
      ? Number(rule.expectedValue || 0)
      : (months.length
        ? avg(months.map((month, monthIndex) => computeRuleMonthlyAmount(rule, month, monthIndex, months, hcMap, grossRevenueMap)))
        : 0);
    baseValues.set(rule.code, normalizedValue);
  });

  const computedValues = new Map();
  const getComputedValue = (code, visiting = new Set()) => {
    const normalizedCode = String(code || "").trim();
    if (computedValues.has(normalizedCode)) return computedValues.get(normalizedCode);
    if (normalizedCode === "A97") return avg(months.map((month) => Number(costMaps.otherRevenueMap[month] || 0)));
    if (normalizedCode === "A91") return avg(months.map((month) => Number(costMaps.insourcingRevenueMap[month] || 0)));
    if (normalizedCode === "A6") return avg(months.map((month) => Number(costMaps.revenueShareMap[month] || 0)));
    if (normalizedCode === "A7") return avg(months.map((month) => Number(costMaps.serverCostMap[month] || 0)));
    if (normalizedCode === "A8") return avg(months.map((month) => Number(costMaps.salesTaxMap[month] || 0)));
    if (normalizedCode === "A54") return avg(months.map((month) => Number(costMaps.royaltyMap[month] || 0)));
    if (normalizedCode === "A9") return avg(months.map((month) => Number(costMaps.grossProfitMap[month] || 0)));
    if (normalizedCode === "A10") return avg(months.map((month) => Number(marketingMaps.salesMarketingMap[month] || 0)));
    if (normalizedCode === "A99") return avg(months.map((month) => Number(marketingMaps.publishingProfitMap[month] || 0)));
    if (!isLinkedRuleCode(normalizedCode)) return Number(baseValues.get(normalizedCode) || 0);
    if (visiting.has(normalizedCode)) return 0;
    visiting.add(normalizedCode);
    const definition = linkedRuleDefinitions[normalizedCode];
    const ctx = {
      netRevenue: months.length ? avg(months.map((month) => Number(netRevenueMap[month] || 0))) : 0,
      adminOtherCodes,
      get: (depCode) => getComputedValue(depCode, new Set(visiting)),
      sum: (codes) => codes.reduce((total, depCode) => total + getComputedValue(depCode, new Set(visiting)), 0)
    };
    const value = Number(definition?.compute(ctx) || 0);
    computedValues.set(normalizedCode, value);
    visiting.delete(normalizedCode);
    return value;
  };

  rules.forEach((rule) => {
    if (!isLinkedRuleCode(rule.code)) return;
    const linkedValue = getComputedValue(rule.code);
    rule.suggestedValue = linkedValue;
    rule.expectedValue = linkedValue;
    rule.expectedValueRaw = String(linkedValue);
    rule.growthRate = 0;
    rule.growthRateRaw = "0";
    rule.type = "fixed";
    rule.method = "fixed";
  });
}

function getActiveScope() {
  return state.importMode === "post" ? "post" : state.ruleScope;
}

function getRulesByScope(scope) {
  return state.rules[scope] || [];
}

function getRowsByScope(scope) {
  return state.actualRows[scope] || [];
}

function getRulesForModel() {
  if (state.importMode === "post") return getRulesByScope("post").filter((rule) => !isManualModuleAccountCode(rule.code));
  return [...getRulesByScope("pre"), ...getRulesByScope("allocated")].filter((rule) => !isManualModuleAccountCode(rule.code));
}

function refreshRules() {
  const scopes = [
    { key: "post", forcedSource: "分摊后" },
    { key: "pre", forcedSource: "项目组直属" },
    { key: "allocated", forcedSource: "分摊值" }
  ];

  scopes.forEach(({ key, forcedSource }) => {
    const nextRules = buildRules(getRowsByScope(key), forcedSource);
    const existing = new Map(getRulesByScope(key).map((rule) => [rule.key, rule]));
    state.rules[key] = nextRules.map((rule) => {
      const previous = existing.get(rule.key);
      const mergedRule = { ...rule, ...(previous || {}) };
      applyRulePreset(mergedRule, { preserveExpectedValue: Boolean(previous) });
      return mergedRule;
    });
    applyLinkedRuleSummaries(key);
  });
  recomputeRuleGridValues();
}

function calculateModel() {
  const months = monthRange(state.forecastStart, state.forecastMonths);
  const revenueSummary = getRevenueForecastSummary(months);
  const grossRevenueMap = Object.fromEntries(revenueSummary.map((row) => [row.month, row.grossRevenue]));
  const vatMap = Object.fromEntries(revenueSummary.map((row) => [row.month, row.vatAmount]));
  const platformMap = Object.fromEntries(revenueSummary.map((row) => [row.month, row.platformAmount]));
  const netRevenueMap = Object.fromEntries(revenueSummary.map((row) => [row.month, row.netRevenue]));
  const costMaps = getCostForecastMaps(months);
  const marketingMaps = getMarketingForecastMaps(months);
  const hcTotalRow = state.hcForecastRows.find((row) => row.code === "A1");
  const hcMap = Object.fromEntries(months.map((month) => [month, Number(hcTotalRow?.values?.[month] || 0)]));
  const trackedRuleCodes = new Set(["A14", "A18", "A53"]);
  const trackedRuleMaps = {
    A14: Object.fromEntries(months.map((month) => [month, 0])),
    A18: Object.fromEntries(months.map((month) => [month, 0])),
    A53: Object.fromEntries(months.map((month) => [month, 0]))
  };

  const monthly = months.map((month) => ({
    month,
    grossRevenue: grossRevenueMap[month],
    vatAmount: vatMap[month],
    platformAmount: platformMap[month],
    netRevenue: netRevenueMap[month],
    otherIncome: 0,
    a97: 0,
    a91: 0,
    a10: 0,
    a11: 0,
    a111: 0,
    a112: 0,
    a12: 0,
    a121: 0,
    a6: 0,
    a7: 0,
    a8: 0,
    a54: 0,
    a14: 0,
    a18: 0,
    a53: 0,
    fixedOtherCost: 0,
    otherOperatingCost: 0,
    totalIncome: 0,
    direct: 0,
    allocated: 0,
    merged: 0,
    profit: 0,
    a94: 0
  }));
  const bySource = {};

  getRulesForModel().forEach((rule) => {
    if (isLinkedRuleCode(rule.code)) return;
    months.forEach((month, index) => {
      const amount = computeRuleMonthlyAmount(rule, month, index, months, hcMap, grossRevenueMap);
      if (!amount) return;
      const row = monthly[index];
      if (modelIncomeRuleCodes.has(rule.code)) {
        row.otherIncome += amount;
      } else if (rule.source === "分摊值") {
        row.allocated += amount;
        bySource[rule.source] = (bySource[rule.source] || 0) + amount;
      } else {
        row.direct += amount;
        bySource[rule.source] = (bySource[rule.source] || 0) + amount;
      }
      if (trackedRuleCodes.has(rule.code)) {
        trackedRuleMaps[rule.code][month] += amount;
      }
    });
  });

  monthly.forEach((row) => {
    const manualCost = Number(costMaps.revenueShareMap[row.month] || 0)
      + Number(costMaps.serverCostMap[row.month] || 0)
      + Number(costMaps.salesTaxMap[row.month] || 0)
      + Number(costMaps.royaltyMap[row.month] || 0)
      + Number(marketingMaps.uaMap[row.month] || 0)
      + Number(marketingMaps.uaMaterialMap[row.month] || 0)
      + Number(marketingMaps.uaTestMap[row.month] || 0)
      + Number(marketingMaps.brandingMap[row.month] || 0)
      + Number(marketingMaps.userOperationMap[row.month] || 0);
    row.a97 = Number(costMaps.otherRevenueMap[row.month] || 0);
    row.a91 = Number(costMaps.insourcingRevenueMap[row.month] || 0);
    row.a10 = Number(marketingMaps.salesMarketingMap[row.month] || 0);
    row.a11 = Number(marketingMaps.uaMap[row.month] || 0);
    row.a111 = Number(marketingMaps.uaMaterialMap[row.month] || 0);
    row.a112 = Number(marketingMaps.uaTestMap[row.month] || 0);
    row.a12 = Number(marketingMaps.brandingMap[row.month] || 0);
    row.a121 = Number(marketingMaps.userOperationMap[row.month] || 0);
    row.a6 = Number(costMaps.revenueShareMap[row.month] || 0);
    row.a7 = Number(costMaps.serverCostMap[row.month] || 0);
    row.a8 = Number(costMaps.salesTaxMap[row.month] || 0);
    row.a54 = Number(costMaps.royaltyMap[row.month] || 0);
    row.a14 = Number(trackedRuleMaps.A14[row.month] || 0);
    row.a18 = Number(trackedRuleMaps.A18[row.month] || 0);
    row.a53 = Number(trackedRuleMaps.A53[row.month] || 0);
    row.otherIncome = row.a97 + row.a91 + row.otherIncome;
    row.merged = row.direct + row.allocated + manualCost;
    row.totalIncome = row.netRevenue + row.otherIncome;
    row.profit = row.netRevenue + row.otherIncome - row.merged;
    row.a94 = row.profit;
    row.fixedOtherCost = row.merged - row.a10 - row.a14 - row.a18 - row.a53 - row.a6 - row.a7 - row.a8 - row.a54;
    row.otherOperatingCost = row.totalIncome - row.profit - row.a10 - row.a14 - row.a18 - row.a53;
  });

  const profitBridge = buildMonthlyProfitBridge(months, getRulesForModel(), hcMap, grossRevenueMap, netRevenueMap, costMaps, marketingMaps);

  return {
    months,
    hcMap,
    revenueSummary,
    monthly,
    profitBridge,
    bySource,
    totalGrossRevenue: sum(monthly.map((row) => row.grossRevenue)),
    totalNetRevenue: sum(monthly.map((row) => row.netRevenue)),
    totalCost: sum(monthly.map((row) => row.merged)),
    totalProfit: sum(monthly.map((row) => row.profit))
  };
}

function getGoLiveMonthKey() {
  return normalizeMonth(state.goLiveDate || state.budgetAlertLaunchMonth || "");
}

function getGoLiveWindowRows(model, limit = 36) {
  const goLiveMonth = getGoLiveMonthKey();
  if (!goLiveMonth) return [...model.monthly];
  return model.monthly.filter((row) => {
    const diff = monthDiff(monthKeyToDate(goLiveMonth), monthKeyToDate(row.month));
    return Number.isFinite(diff) && diff >= 0 && diff < limit;
  });
}

function getPreGoLiveRows(model) {
  const goLiveMonth = getGoLiveMonthKey();
  if (!goLiveMonth) return [];
  return model.monthly.filter((row) => row.month < goLiveMonth);
}

function getStaticPaybackMonthsFromRows(rows, goLiveMonth = getGoLiveMonthKey()) {
  if (!goLiveMonth) return null;
  let runningProfit = 0;
  for (const row of rows) {
    runningProfit += Number(row.profit || 0);
    if (row.month >= goLiveMonth && runningProfit >= 0) {
      const diff = monthDiff(monthKeyToDate(goLiveMonth), monthKeyToDate(row.month));
      return Number.isFinite(diff) ? diff + 1 : null;
    }
  }
  return null;
}

function formatPaybackMonths(value) {
  return Number.isFinite(value) && value > 0 ? `${value} 个月` : "未回本";
}

function formatBoardMoney(value, digits = 1) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    notation: "compact",
    compactDisplay: "short",
    minimumFractionDigits: 0,
    maximumFractionDigits: digits
  }).format(Number(value || 0));
}

function formatGlMonthIndex(month) {
  const goLiveMonth = getGoLiveMonthKey();
  if (!goLiveMonth) return formatMonthSlashLabel(month);
  const diff = monthDiff(monthKeyToDate(goLiveMonth), monthKeyToDate(month));
  if (!Number.isFinite(diff)) return formatMonthSlashLabel(month);
  return diff >= 0 ? `GL${diff}M` : `研发${diff}M`;
}

function getGlCoverageLabel(overview, model) {
  if (!overview.goLiveMonth) {
    return `当前预测区间：${formatMonthSlashLabel(model.months[0])} - ${formatMonthSlashLabel(model.months.at(-1))}`;
  }
  if (!overview.glCoverageMonths) {
    return `已填写 GL 时间：${formatMonthSlashLabel(overview.goLiveMonth)}；当前预测区间尚未覆盖 GL 后月份`;
  }
  const lastGlIndex = overview.glCoverageMonths - 1;
  if (overview.glCoverageMonths >= 36) {
    return "统计区间：GL3Y（GL0M - GL35M）";
  }
  return `统计区间：GL后已覆盖 ${overview.glCoverageMonths} 个月（GL0M - GL${lastGlIndex}M）`;
}

function buildFinancialOverview(model) {
  const glRows = getGoLiveWindowRows(model, 36);
  const preGlRows = getPreGoLiveRows(model);
  const glGrossRevenue = sum(glRows.map((row) => Number(row.grossRevenue || 0)));
  const glNetProfit = sum(glRows.map((row) => Number(row.a94 || 0)));
  const glCost = sum(glRows.map((row) => Number(row.merged || 0)));
  const glMargin = glGrossRevenue ? glNetProfit / glGrossRevenue : 0;
  const totalRoi = model.totalCost ? model.totalProfit / model.totalCost : 0;
  const peakRevenueRow = glRows.reduce((best, row) => (
    Number(row.grossRevenue || 0) > Number(best?.grossRevenue || 0) ? row : best
  ), glRows[0] || null);
  const peakCostRow = model.monthly.reduce((best, row) => (
    Number(row.merged || 0) > Number(best?.merged || 0) ? row : best
  ), model.monthly[0] || null);
  const paybackMonths = getStaticPaybackMonthsFromRows(model.monthly);

  return {
    goLiveMonth: getGoLiveMonthKey(),
    glRows,
    preGlRows,
    glCoverageMonths: glRows.length,
    glGrossRevenue,
    glNetProfit,
    glCost,
    glMargin,
    totalRoi,
    paybackMonths,
    peakRevenueRow,
    peakCostRow
  };
}

function formatResultAxisLabel(month) {
  const goLiveMonth = getGoLiveMonthKey();
  if (!goLiveMonth) return formatMonthSlashLabel(month);
  const diff = monthDiff(monthKeyToDate(goLiveMonth), monthKeyToDate(month));
  if (!Number.isFinite(diff)) return formatMonthSlashLabel(month);
  if (diff < 0) return `研发${diff}`;
  if (diff === 0) return "GL首月";
  return `GL${diff}M`;
}

function buildFlowMixSeries(model, dimension = state.resultRevenueView) {
  const targetRows = getGoLiveWindowRows(model, 36);
  const breakdownKey = dimension === "platform" ? "platformBreakdown" : "regionBreakdown";
  const defaultOrder = dimension === "platform" ? ["移动", "PC", "Console"] : ["CN", "海外"];
  const rowsWithBreakdown = state.flowRows.filter((row) => Object.keys(row?.[breakdownKey] || {}).length);

  if (!rowsWithBreakdown.length || !targetRows.length) {
    return { labels: [], months: targetRows.map((row) => row.month), series: [] };
  }

  const averageBreakdown = {};
  let averageTotal = 0;
  rowsWithBreakdown.forEach((row) => {
    const rowTotal = sum(Object.values(row[breakdownKey] || {}).map((value) => Number(value || 0)));
    if (!rowTotal) return;
    averageTotal += rowTotal;
    Object.entries(row[breakdownKey] || {}).forEach(([label, value]) => {
      averageBreakdown[label] = Number(averageBreakdown[label] || 0) + Number(value || 0);
    });
  });

  const fallbackShares = Object.fromEntries(
    Object.entries(averageBreakdown).map(([label, value]) => [label, averageTotal ? Number(value || 0) / averageTotal : 0])
  );

  const allLabels = [...new Set([
    ...defaultOrder,
    ...rowsWithBreakdown.flatMap((row) => Object.keys(row[breakdownKey] || {}))
  ].filter(Boolean))];

  const flowMonthMap = new Map(rowsWithBreakdown.map((row) => {
    const rowTotal = sum(Object.values(row[breakdownKey] || {}).map((value) => Number(value || 0)));
    const shares = Object.fromEntries(
      Object.entries(row[breakdownKey] || {}).map(([label, value]) => [label, rowTotal ? Number(value || 0) / rowTotal : 0])
    );
    return [row.month, shares];
  }));

  const series = allLabels.map((label) => ({
    name: label,
    data: targetRows.map((row) => {
      const monthShares = flowMonthMap.get(row.month) || fallbackShares;
      return Math.max(Number(row.grossRevenue || 0) * Number(monthShares?.[label] || 0), 0);
    })
  }));

  return {
    labels: allLabels,
    months: targetRows.map((row) => row.month),
    series
  };
}

function buildFinancialStressRows(model, { revenueDownRate = 0, uaUpRate = 0 } = {}) {
  const revenueFactor = 1 - revenueDownRate;
  const uaFactor = 1 + uaUpRate;

  return model.monthly.map((row) => {
    const nextGrossRevenue = Number(row.grossRevenue || 0) * revenueFactor;
    const nextVatAmount = Number(row.vatAmount || 0) * revenueFactor;
    const nextPlatformAmount = Number(row.platformAmount || 0) * revenueFactor;
    const nextNetRevenue = Number(row.netRevenue || 0) * revenueFactor;
    const nextA97 = Number(row.a97 || 0) * revenueFactor;
    const nextA91 = Number(row.a91 || 0) * revenueFactor;
    const nextA6 = Number(row.a6 || 0) * revenueFactor;
    const nextA7 = Number(row.a7 || 0) * revenueFactor;
    const nextA8 = Number(row.a8 || 0) * revenueFactor;
    const nextA54 = Number(row.a54 || 0) * revenueFactor;
    const nextA11 = Number(row.a11 || 0) * revenueFactor * uaFactor;
    const nextA111 = Number(row.a111 || 0) * revenueFactor * uaFactor;
    const nextA112 = Number(row.a112 || 0) * revenueFactor * uaFactor;
    const nextA12 = Number(row.a12 || 0);
    const nextA121 = Number(row.a121 || 0);
    const nextA10 = nextA11 + nextA111 + nextA112 + nextA12 + nextA121;
    const nextMerged = Number(row.a14 || 0)
      + Number(row.a18 || 0)
      + Number(row.a53 || 0)
      + nextA10
      + nextA6
      + nextA7
      + nextA8
      + nextA54
      + Number(row.fixedOtherCost || 0);
    const nextOtherIncome = nextA97 + nextA91;
    const nextProfit = nextNetRevenue + nextOtherIncome - nextMerged;
    return {
      ...row,
      grossRevenue: nextGrossRevenue,
      vatAmount: nextVatAmount,
      platformAmount: nextPlatformAmount,
      netRevenue: nextNetRevenue,
      a97: nextA97,
      a91: nextA91,
      a6: nextA6,
      a7: nextA7,
      a8: nextA8,
      a54: nextA54,
      a11: nextA11,
      a111: nextA111,
      a112: nextA112,
      a10: nextA10,
      otherIncome: nextOtherIncome,
      totalIncome: nextNetRevenue + nextOtherIncome,
      merged: nextMerged,
      otherOperatingCost: nextA6 + nextA7 + nextA8 + nextA54 + Number(row.fixedOtherCost || 0),
      profit: nextProfit,
      a94: nextProfit
    };
  });
}

function buildStressScenarioSummary(rows) {
  const glRows = getGoLiveMonthKey()
    ? rows.filter((row) => {
        const diff = monthDiff(monthKeyToDate(getGoLiveMonthKey()), monthKeyToDate(row.month));
        return Number.isFinite(diff) && diff >= 0 && diff < 36;
      })
    : rows;
  const glGrossRevenue = sum(glRows.map((row) => Number(row.grossRevenue || 0)));
  const glProfit = sum(glRows.map((row) => Number(row.profit || 0)));
  return {
    paybackMonths: getStaticPaybackMonthsFromRows(rows),
    glProfit,
    margin: glGrossRevenue ? glProfit / glGrossRevenue : 0
  };
}

function buildStressScenarioSet(model) {
  const baselineRows = model.monthly.map((row) => ({ ...row }));
  const revenueRows = buildFinancialStressRows(model, { revenueDownRate: state.stressRevenueDownRate, uaUpRate: 0 });
  const uaRows = buildFinancialStressRows(model, { revenueDownRate: 0, uaUpRate: state.stressUaUpRate });

  return {
    baselineRows,
    revenueRows,
    uaRows,
    baseline: buildStressScenarioSummary(baselineRows),
    revenue: buildStressScenarioSummary(revenueRows),
    ua: buildStressScenarioSummary(uaRows)
  };
}

function createSummaryAggregate(label, months = []) {
  return {
    label,
    months: [...months],
    grossRevenue: 0,
    vatAmount: 0,
    platformAmount: 0,
    netRevenue: 0,
    merged: 0,
    profit: 0
  };
}

function appendSummaryRow(target, row) {
  target.grossRevenue += Number(row.grossRevenue || 0);
  target.vatAmount += Number(row.vatAmount || 0);
  target.platformAmount += Number(row.platformAmount || 0);
  target.netRevenue += Number(row.netRevenue || 0);
  target.merged += Number(row.merged || 0);
  target.profit += Number(row.profit || 0);
}

function formatCoverageMonths(months) {
  if (!months.length) return "-";
  if (months.length === 1) return formatMonthSlashLabel(months[0]);
  return `${formatMonthSlashLabel(months[0])} - ${formatMonthSlashLabel(months.at(-1))}`;
}

function buildNaturalYearSummaries(model) {
  const yearMap = new Map();
  model.monthly.forEach((row) => {
    const year = String(row.month || "").slice(0, 4) || "-";
    if (!yearMap.has(year)) yearMap.set(year, createSummaryAggregate(year));
    const target = yearMap.get(year);
    target.months.push(row.month);
    appendSummaryRow(target, row);
  });
  return [...yearMap.values()];
}

function buildGoLiveSummaries(model, goLiveDate) {
  const goLive = parseInputDate(goLiveDate);
  const buckets = Array.from({ length: 5 }, (_, index) => createSummaryAggregate(`GL${index + 1}Y`));
  if (!goLive) return buckets;

  const goLiveMonth = new Date(goLive.getFullYear(), goLive.getMonth(), 1);
  model.monthly.forEach((row) => {
    const monthDate = monthKeyToDate(row.month);
    const diff = monthDiff(goLiveMonth, monthDate);
    if (!Number.isFinite(diff) || diff < 0 || diff >= 60) return;
    const bucketIndex = Math.floor(diff / 12);
    const target = buckets[bucketIndex];
    target.months.push(row.month);
    appendSummaryRow(target, row);
  });
  return buckets;
}

function renderAnnualSummaryRows(rows, emptyText, { keepEmptyRows = false } = {}) {
  const targetRows = keepEmptyRows ? rows : rows.filter((row) => row.months.length);
  if (!targetRows.length) {
    return `<tr><td colspan="8" class="table-note">${emptyText}</td></tr>`;
  }
  return targetRows.map((row) => `
    <tr>
      <td>${row.label}</td>
      <td>${formatCoverageMonths(row.months)}</td>
      <td>${row.months.length ? compactMoney(row.grossRevenue) : "-"}</td>
      <td>${row.months.length ? compactMoney(row.vatAmount) : "-"}</td>
      <td>${row.months.length ? compactMoney(row.platformAmount) : "-"}</td>
      <td>${row.months.length ? compactMoney(row.netRevenue) : "-"}</td>
      <td>${row.months.length ? compactMoney(row.merged) : "-"}</td>
      <td class="${row.profit >= 0 ? "positive" : "negative"}">${row.months.length ? compactMoney(row.profit) : "-"}</td>
    </tr>
  `).join("");
}

const financialReportLineDefinitions = [
  { type: "section", label: "收入链路" },
  { code: "A2", label: "A2 项目流水", getter: (row) => row.grossRevenue },
  { code: "A3", label: "A3 增值税", getter: (row) => row.vatAmount, tone: "negative" },
  { code: "A4", label: "A4 平台费", getter: (row) => row.platformAmount, tone: "negative" },
  { code: "A5", label: "A5 净收入", getter: (row) => row.netRevenue },
  { code: "A97", label: "A97 其他收入", getter: (row) => row.a97 },
  { code: "A91", label: "A91 内采收入", getter: (row) => row.a91 },
  { code: "A9", label: "A9 毛利润", getter: (row) => row.netRevenue + row.a97 + row.a91 - row.a6 - row.a7 - row.a8 - row.a54 },
  { type: "section", label: "市场与发行" },
  { code: "A10", label: "A10 市场费用", getter: (row) => row.a10, tone: "negative" },
  { code: "A11", label: "A11 买量投放", getter: (row) => row.a11, indent: 1, tone: "negative" },
  { code: "A11.1", label: "A11.1 素材制作", getter: (row) => row.a111, indent: 1, tone: "negative" },
  { code: "A11.2", label: "A11.2 投放测试", getter: (row) => row.a112, indent: 1, tone: "negative" },
  { code: "A12", label: "A12 品牌营销", getter: (row) => row.a12, indent: 1, tone: "negative" },
  { code: "A12.1", label: "A12.1 用户运营", getter: (row) => row.a121, indent: 1, tone: "negative" },
  { code: "A99", label: "A99 发行利润", getter: (row) => row.netRevenue + row.a97 + row.a91 - row.a6 - row.a7 - row.a8 - row.a54 - row.a10 },
  { type: "section", label: "研发与其他成本" },
  { code: "A14", label: "A14 研发人力", getter: (row) => row.a14, tone: "negative" },
  { code: "A18", label: "A18 外包费用", getter: (row) => row.a18, tone: "negative" },
  { code: "A53", label: "A53 内包费用", getter: (row) => row.a53, tone: "negative" },
  { code: "A6", label: "A6 分成成本", getter: (row) => row.a6, tone: "negative" },
  { code: "A7", label: "A7 服务器成本", getter: (row) => row.a7, tone: "negative" },
  { code: "A8", label: "A8 营业税", getter: (row) => row.a8, tone: "negative" },
  { code: "A54", label: "A54 版权金", getter: (row) => row.a54, tone: "negative" },
  { code: "OTH", label: "其他经营成本", getter: (row) => row.fixedOtherCost || 0, tone: "negative" },
  { type: "section", label: "最终利润" },
  { code: "A94", label: "A94 项目利润", getter: (row) => row.a94, tone: "profit" }
];

function createFinancialReportBucket(label, meta = {}) {
  return {
    label,
    meta,
    rows: []
  };
}

function buildNaturalYearReportColumns(model) {
  const yearMap = new Map();
  model.monthly.forEach((row) => {
    const year = String(row.month || "").slice(0, 4) || "-";
    if (!yearMap.has(year)) yearMap.set(year, createFinancialReportBucket(year));
    yearMap.get(year).rows.push(row);
  });
  return [...yearMap.values()].map((column) => ({
    ...column,
    meta: {
      coverage: formatCoverageMonths(column.rows.map((row) => row.month))
    }
  }));
}

function buildGoLiveReportColumns(model, goLiveDate) {
  const effectiveGoLive = normalizeMonth(goLiveDate || state.budgetAlertLaunchMonth || "");
  const goLive = parseInputDate(effectiveGoLive ? `${effectiveGoLive}-01` : "");
  const buckets = Array.from({ length: 5 }, (_, index) => createFinancialReportBucket(`GL${index + 1}Y`));
  if (!goLive) return [];

  const goLiveMonth = new Date(goLive.getFullYear(), goLive.getMonth(), 1);
  model.monthly.forEach((row) => {
    const monthDate = monthKeyToDate(row.month);
    const diff = monthDiff(goLiveMonth, monthDate);
    if (!Number.isFinite(diff) || diff < 0 || diff >= 60) return;
    const bucketIndex = Math.floor(diff / 12);
    buckets[bucketIndex].rows.push(row);
  });
  return buckets.map((bucket, index) => ({
    ...bucket,
    meta: {
      coverage: bucket.rows.length
        ? `${formatGlMonthIndex(bucket.rows[0].month)} - ${formatGlMonthIndex(bucket.rows.at(-1).month)}`
        : `GL${index + 1}Y 未覆盖`
    }
  })).filter((bucket) => bucket.rows.length);
}

function renderFinancialReportTable(columns, emptyText) {
  if (!columns.length) {
    return `<div class="empty-state">${emptyText}</div>`;
  }
  const targetColumns = columns;

  const thead = `
    <thead>
      <tr>
        <th class="report-item-col">科目</th>
        ${targetColumns.map((column) => `<th>${column.label}<span class="report-col-sub">${escapeHtml(column.meta?.coverage || "-")}</span></th>`).join("")}
      </tr>
    </thead>
  `;

  const tbody = `
    <tbody>
      ${financialReportLineDefinitions.map((item) => {
        if (item.type === "section") {
          return `<tr class="report-section-row"><th class="report-item-col" colspan="${targetColumns.length + 1}">${item.label}</th></tr>`;
        }
        const values = targetColumns.map((column) => sum(column.rows.map((row) => Number(item.getter(row) || 0))));
        return `
          <tr class="report-data-row ${item.tone === "profit" ? "report-profit-row" : ""}">
            <th class="report-item-col report-item-indent-${item.indent || 0}">
              <span>${item.label}</span>
            </th>
            ${values.map((value) => `<td class="${item.tone === "profit" ? "positive" : item.tone === "negative" ? "negative" : ""}">${formatBoardMoney(value)}</td>`).join("")}
          </tr>
        `;
      }).join("")}
    </tbody>
  `;

  return `<table class="finance-report-table">${thead}${tbody}</table>`;
}

function renderPeriodSummaryTables(model) {
  if (els.annualReportWrap) {
    const annualColumns = buildNaturalYearReportColumns(model);
    els.annualReportWrap.innerHTML = renderFinancialReportTable(
      annualColumns,
      "当前预测区间内暂无可展示的自然年报表。"
    );
  }
  if (els.glReportWrap) {
    const glColumns = buildGoLiveReportColumns(model, state.goLiveDate);
    els.glReportWrap.innerHTML = renderFinancialReportTable(
      glColumns,
      getGoLiveMonthKey() ? "当前预测区间尚未覆盖 GL 年报表。" : "请先在步骤 2.2 填写 GL 时间；若未填写，则回退读取步骤 1 的预计上线月。"
    );
  }
}

function ensureChartInstance(host, currentInstance) {
  if (!host || !window.echarts) return null;
  if (currentInstance && !currentInstance.isDisposed?.()) return currentInstance;
  return window.echarts.getInstanceByDom(host) || window.echarts.init(host, null, { renderer: "canvas" });
}

function renderMetrics(model) {
  const overview = buildFinancialOverview(model);
  const grossTitle = overview.goLiveMonth && overview.glCoverageMonths >= 36
    ? "GL3Y累计总流水(A2)"
    : "GL后累计总流水(A2)";
  const profitTitle = overview.goLiveMonth && overview.glCoverageMonths >= 36
    ? "GL3Y累计项目利润(A94)"
    : "GL后累计项目利润(A94)";
  const coverageLabel = getGlCoverageLabel(overview, model);
  const peakRevenueNote = overview.peakRevenueRow
    ? `峰值月：${formatGlMonthIndex(overview.peakRevenueRow.month)}（${formatMonthSlashLabel(overview.peakRevenueRow.month)}） / ${formatBoardMoney(overview.peakRevenueRow.grossRevenue)}`
    : "请先填写 GL 时间后查看 GL 口径峰值月。";
  const metrics = [
    {
      title: grossTitle,
      value: formatBoardMoney(overview.glGrossRevenue || model.totalGrossRevenue),
      tone: "",
      note: `${coverageLabel}<br/>${peakRevenueNote}`
    },
    {
      title: profitTitle,
      value: formatBoardMoney(overview.glNetProfit || model.totalProfit),
      tone: (overview.glNetProfit || model.totalProfit) >= 0 ? "positive" : "negative",
      note: `${coverageLabel}<br/>项目利润率：${ratio(overview.glMargin)}，定义为 GL 口径项目利润 / GL 口径项目流水。`
    },
    {
      title: "GL后累计盈利时间",
      value: formatPaybackMonths(overview.paybackMonths),
      tone: overview.paybackMonths ? "positive" : "negative",
      note: `定义：从 GL0M 起累计项目利润首次转正的时间。<br/>全盘 ROI：${ratio(overview.totalRoi)}，定义为全周期累计利润 / 全周期累计成本。`
    }
  ];

  els.metricGrid.innerHTML = metrics.map((item) => `
    <article class="metric metric-finance-card">
      <div class="metric-title">${item.title}</div>
      <div class="metric-value ${item.tone}">${item.value}</div>
      <div class="metric-note">${item.note}</div>
    </article>
  `).join("");
}

function renderBudgetStrip(model) {
  if (!els.resultBudgetStrip) return;
  if (!state.projectBudgetAmount) {
    els.resultBudgetStrip.innerHTML = `
      <div class="result-budget-pill">
        <strong>预算提示</strong>
        <span>步骤 1 尚未填写项目预算，当前步骤 3 仅展示利润与结构分析，不展示预算偏差。</span>
      </div>
    `;
    return;
  }

  const budgetGap = state.projectBudgetAmount - model.totalCost;
  const budgetUsage = state.projectBudgetAmount ? model.totalCost / state.projectBudgetAmount : 0;
  const statusClass = budgetGap >= 0 ? "is-positive" : "is-negative";
  const statusText = budgetGap >= 0 ? `预算结余 ${compactMoney(budgetGap)}` : `预算超支 ${compactMoney(Math.abs(budgetGap))}`;

  els.resultBudgetStrip.innerHTML = `
    <div class="result-budget-pill ${statusClass}">
      <strong>预算执行</strong>
      <span>累计合并成本 ${compactMoney(model.totalCost)} / 项目预算 ${compactMoney(state.projectBudgetAmount)}，执行进度 ${ratio(budgetUsage)}，${statusText}。</span>
    </div>
  `;
}

function renderSummaryTable(model) {
  if (!els.summaryBody) return;
  els.summaryBody.innerHTML = model.monthly.map((row) => `
    <tr>
      <td>${row.month}</td>
      <td>${compactMoney(row.grossRevenue)}</td>
      <td>${formatMoneyWithRatio(row.vatAmount, row.grossRevenue ? row.vatAmount / row.grossRevenue : 0)}</td>
      <td>${formatMoneyWithRatio(row.platformAmount, row.grossRevenue ? row.platformAmount / row.grossRevenue : 0)}</td>
      <td>${formatMoneyWithRatio(row.netRevenue, row.grossRevenue ? row.netRevenue / row.grossRevenue : 0)}</td>
      <td>${compactMoney(row.direct)}</td>
      <td>${compactMoney(row.allocated)}</td>
      <td>${compactMoney(row.merged)}</td>
      <td class="${row.profit >= 0 ? "positive" : "negative"}">${compactMoney(row.profit)}</td>
    </tr>
  `).join("");
}

function renderProfitBridgeTable(model) {
  if (!els.profitBridgeBody) return;
  els.profitBridgeBody.innerHTML = model.profitBridge.map((row, index) => `
    <tr>
      <td>${row.month}</td>
      <td>${compactMoney(row.a5)}</td>
      <td>${compactMoney(row.a97)}</td>
      <td>${compactMoney(row.a91)}</td>
      <td>${compactMoney(row.a9)}</td>
      <td>${compactMoney(row.a10)}</td>
      <td>${compactMoney(row.a99)}</td>
      <td>${compactMoney(row.a14)}</td>
      <td>${compactMoney(row.a18)}</td>
      <td>${compactMoney(row.a13)}</td>
      <td class="${model.monthly[index]?.profit >= 0 ? "positive" : "negative"}">${compactMoney(model.monthly[index]?.profit || 0)}</td>
    </tr>
  `).join("");
}

function renderProfitBridgeChart(model) {
  if (!els.profitBridgeChart) return;
  const series = [
    { key: "a9", label: "A9 毛利润", color: "#0f6cbd" },
    { key: "a99", label: "A99 发行利润", color: "#f77f00" },
    { key: "profit", label: "项目利润", color: "#1f8f6a" }
  ];
  const maxValue = Math.max(
    ...model.profitBridge.flatMap((row, index) => [Math.abs(row.a9), Math.abs(row.a99), Math.abs(model.monthly[index]?.profit || 0)]),
    1
  );
  const baseY = 240;
  const scale = 150 / maxValue;
  const pointX = (index) => 70 + index * (760 / Math.max(model.profitBridge.length - 1, 1));
  const pointY = (value) => baseY - value * scale;
  const pathFor = (getter) => model.profitBridge.map((row, index) => `${index === 0 ? "M" : "L"} ${pointX(index)} ${pointY(getter(row, index))}`).join(" ");

  const paths = [
    pathFor((row) => row.a9),
    pathFor((row) => row.a99),
    pathFor((row, index) => model.monthly[index]?.profit || 0)
  ];

  const labels = model.profitBridge.map((row, index) => `
    <text x="${pointX(index)}" y="274" font-size="11" text-anchor="middle" fill="#64748b">${row.month.slice(5)}</text>
  `).join("");

  const dots = series.map((item) => model.profitBridge.map((row, index) => {
    const value = item.key === "profit" ? (model.monthly[index]?.profit || 0) : row[item.key];
    return `<circle cx="${pointX(index)}" cy="${pointY(value)}" r="4" fill="${item.color}"></circle>`;
  }).join("")).join("");

  const legend = series.map((item, index) => `
    <circle cx="${70 + index * 150}" cy="26" r="5" fill="${item.color}"></circle>
    <text x="${82 + index * 150}" y="30" font-size="12" fill="${item.color}">${item.label}</text>
  `).join("");

  els.profitBridgeChart.innerHTML = `
    <line x1="50" y1="${baseY}" x2="870" y2="${baseY}" stroke="rgba(19,34,56,0.18)"></line>
    <line x1="50" y1="90" x2="870" y2="90" stroke="rgba(19,34,56,0.08)" stroke-dasharray="4 4"></line>
    <line x1="50" y1="${pointY(-maxValue)}" x2="870" y2="${pointY(-maxValue)}" stroke="rgba(19,34,56,0.08)" stroke-dasharray="4 4"></line>
    ${legend}
    <path d="${paths[0]}" fill="none" stroke="${series[0].color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
    <path d="${paths[1]}" fill="none" stroke="${series[1].color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
    <path d="${paths[2]}" fill="none" stroke="${series[2].color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
    ${dots}
    ${labels}
  `;
}

function renderWaterfallChart(model) {
  waterfallChartInstance = ensureChartInstance(els.waterfallChart, waterfallChartInstance);
  if (!waterfallChartInstance) return;

  const overview = buildFinancialOverview(model);
  const targetRows = overview.glRows.length ? overview.glRows : model.monthly;
  const grossRevenue = sum(targetRows.map((row) => Number(row.grossRevenue || 0)));
  const vatAmount = sum(targetRows.map((row) => Number(row.vatAmount || 0)));
  const platformAmount = sum(targetRows.map((row) => Number(row.platformAmount || 0)));
  const otherIncome = sum(targetRows.map((row) => Number(row.a97 || 0) + Number(row.a91 || 0)));
  const marketingCost = sum(targetRows.map((row) => Number(row.a10 || 0)));
  const productionCost = sum(targetRows.map((row) => Number(row.a14 || 0) + Number(row.a18 || 0) + Number(row.a53 || 0)));
  const otherOperatingCost = sum(targetRows.map((row) => Number(row.otherOperatingCost || 0)));
  const profit = sum(targetRows.map((row) => Number(row.a94 || 0)));

  const categories = ["A2 项目流水", "A3 增值税", "A4 平台费", "其他经营收入", "A10 市场费用", "研发人力与制作", "其他经营成本", "A94 项目利润"];
  const deltas = [-vatAmount, -platformAmount, otherIncome, -marketingCost, -productionCost, -otherOperatingCost];
  const placeholder = ["-", ...deltas.map(() => 0), "-"];
  const changeSeries = ["-", ...deltas.map(() => 0), "-"];
  let running = grossRevenue;
  deltas.forEach((delta, index) => {
    if (delta >= 0) {
      placeholder[index + 1] = running;
      changeSeries[index + 1] = delta;
    } else {
      placeholder[index + 1] = running + delta;
      changeSeries[index + 1] = Math.abs(delta);
    }
    running += delta;
  });

  const totalSeries = [grossRevenue, "-", "-", "-", "-", "-", "-", profit];
  const itemStyles = [
    { color: "#2563eb" },
    { color: "#ef4444" },
    { color: "#f97316" },
    { color: "#14b8a6" },
    { color: "#f43f5e" },
    { color: "#6366f1" },
    { color: "#475569" },
    { color: profit >= 0 ? "#16a34a" : "#dc2626" }
  ];

  waterfallChartInstance.setOption({
    animation: false,
    tooltip: {
      trigger: "axis",
      axisPointer: { type: "shadow" },
      formatter(params) {
        const current = params.find((item) => item.seriesName !== "占位");
        if (!current) return "";
        const category = current.axisValueLabel;
        const rawValue = category === "A2 项目流水"
          ? grossRevenue
          : category === "A94 项目利润"
            ? profit
            : deltas[categories.indexOf(category) - 1];
        return `${category}<br/>${compactMoney(rawValue)}`;
      }
    },
    grid: { left: "3%", right: "3%", top: 16, bottom: 24, containLabel: true },
    xAxis: {
      type: "category",
      data: categories,
      axisLabel: { interval: 0, fontSize: 11, color: "#64748b" },
      axisLine: { lineStyle: { color: "#cbd5e1" } }
    },
    yAxis: {
      type: "value",
      axisLabel: { color: "#94a3b8", formatter: (value) => compactMoney(value) },
      splitLine: { lineStyle: { color: "#e2e8f0", type: "dashed" } }
    },
    series: [
      {
        name: "占位",
        type: "bar",
        stack: "总量",
        silent: true,
        itemStyle: { color: "transparent", borderColor: "transparent" },
        emphasis: { itemStyle: { color: "transparent", borderColor: "transparent" } },
        data: placeholder
      },
      {
        name: "变化项",
        type: "bar",
        stack: "总量",
        barMaxWidth: 38,
        label: { show: true, position: "insideTop", color: "#0f172a", formatter: ({ dataIndex }) => dataIndex > 0 && dataIndex < categories.length - 1 ? compactMoney(deltas[dataIndex - 1]) : "" },
        itemStyle: { borderRadius: 10, color: ({ dataIndex }) => itemStyles[dataIndex].color },
        data: changeSeries
      },
      {
        name: "总计",
        type: "bar",
        barMaxWidth: 38,
        label: { show: true, position: "top", color: "#0f172a", formatter: ({ value }) => value === "-" ? "" : compactMoney(value) },
        itemStyle: { borderRadius: 10, color: ({ dataIndex }) => itemStyles[dataIndex].color },
        data: totalSeries
      }
    ]
  }, true);
}

function renderRevenueStructureChart(model) {
  if (els.revenueViewToggle) {
    [...els.revenueViewToggle.querySelectorAll("[data-revenue-view]")].forEach((button) => {
      button.classList.toggle("is-active", button.dataset.revenueView === state.resultRevenueView);
    });
  }
  revenueStructureChartInstance = ensureChartInstance(els.revenueStructureChart, revenueStructureChartInstance);
  if (!revenueStructureChartInstance) return;

  const seriesModel = buildFlowMixSeries(model, state.resultRevenueView);
  if (!seriesModel.series.length) {
    revenueStructureChartInstance.clear();
    if (els.revenueStructureChart) {
      els.revenueStructureChart.innerHTML = `<div class="empty-state">当前流水明细中还未识别到${state.resultRevenueView === "platform" ? "平台" : "地区"}字段，请导入包含细分维度的流水 Excel 后查看。</div>`;
    }
    return;
  }

  const palette = state.resultRevenueView === "platform"
    ? ["#3b82f6", "#8b5cf6", "#f59e0b"]
    : ["#2563eb", "#10b981", "#8b5cf6", "#f97316"];
  if (els.revenueStructureChart?.querySelector(".empty-state")) {
    els.revenueStructureChart.innerHTML = "";
    revenueStructureChartInstance = ensureChartInstance(els.revenueStructureChart, null);
  }

  revenueStructureChartInstance.setOption({
    animation: false,
    tooltip: { trigger: "axis", axisPointer: { type: "cross" } },
    legend: { top: 0, icon: "circle", textStyle: { color: "#475569" } },
    grid: { left: "3%", right: "4%", top: 36, bottom: 20, containLabel: true },
    xAxis: {
      type: "category",
      boundaryGap: false,
      data: seriesModel.months.map((month) => formatResultAxisLabel(month)),
      axisLabel: { interval: Math.max(Math.floor(seriesModel.months.length / 8), 0), color: "#94a3b8" },
      axisLine: { lineStyle: { color: "#cbd5e1" } }
    },
    yAxis: {
      type: "value",
      axisLabel: { color: "#94a3b8", formatter: (value) => compactMoney(value) },
      splitLine: { lineStyle: { color: "#e2e8f0", type: "dashed" } }
    },
    series: seriesModel.series.map((item, index) => ({
      name: item.name,
      type: "line",
      stack: "total",
      smooth: true,
      symbol: "none",
      areaStyle: { opacity: 0.18 },
      emphasis: { focus: "series" },
      color: palette[index % palette.length],
      data: item.data
    }))
  }, true);
}

function renderCostStructureChart(model) {
  costStructureChartInstance = ensureChartInstance(els.costStructureChart, costStructureChartInstance);
  if (!costStructureChartInstance) return;

  costStructureChartInstance.setOption({
    animation: false,
    tooltip: { trigger: "axis", axisPointer: { type: "cross" } },
    legend: { top: 0, icon: "roundRect", textStyle: { color: "#475569" } },
    grid: { left: "3%", right: "4%", top: 36, bottom: 62, containLabel: true },
    xAxis: {
      type: "category",
      boundaryGap: false,
      data: model.monthly.map((row) => formatResultAxisLabel(row.month)),
      axisLabel: { interval: 0, rotate: 42, color: "#94a3b8", fontSize: 10 },
      axisLine: { lineStyle: { color: "#cbd5e1" } }
    },
    yAxis: {
      type: "value",
      axisLabel: { color: "#94a3b8", formatter: (value) => compactMoney(value) },
      splitLine: { lineStyle: { color: "#e2e8f0", type: "dashed" } }
    },
    series: [
      {
        name: "研发人力(A14)",
        type: "line",
        stack: "cost",
        smooth: true,
        symbol: "none",
        areaStyle: { opacity: 0.2 },
        color: "#6366f1",
        data: model.monthly.map((row) => Number(row.a14 || 0))
      },
      {
        name: "内包与外包(A18+A53)",
        type: "line",
        stack: "cost",
        smooth: true,
        symbol: "none",
        areaStyle: { opacity: 0.18 },
        color: "#f59e0b",
        data: model.monthly.map((row) => Number(row.a18 || 0) + Number(row.a53 || 0))
      },
      {
        name: "营销与买量(A10)",
        type: "line",
        stack: "cost",
        smooth: true,
        symbol: "none",
        areaStyle: { opacity: 0.18 },
        color: "#ef4444",
        data: model.monthly.map((row) => Number(row.a10 || 0))
      },
      {
        name: "其他经营成本",
        type: "line",
        stack: "cost",
        smooth: true,
        symbol: "none",
        areaStyle: { opacity: 0.16 },
        color: "#475569",
        data: model.monthly.map((row) => Number(row.otherOperatingCost || 0))
      }
    ]
  }, true);
}

function renderStressTest(model) {
  if (els.stressRevenueDownRange) els.stressRevenueDownRange.value = String(Math.round(state.stressRevenueDownRate * 100));
  if (els.stressUaUpRange) els.stressUaUpRange.value = String(Math.round(state.stressUaUpRate * 100));
  if (els.stressRevenueDownValue) els.stressRevenueDownValue.textContent = `${Math.round(state.stressRevenueDownRate * 100)}%`;
  if (els.stressUaUpValue) els.stressUaUpValue.textContent = `${Math.round(state.stressUaUpRate * 100)}%`;

  const scenarioSet = buildStressScenarioSet(model);
  if (els.stressRevenuePayback) els.stressRevenuePayback.textContent = formatPaybackMonths(scenarioSet.revenue.paybackMonths);
  if (els.stressRevenueProfit) els.stressRevenueProfit.textContent = compactMoney(scenarioSet.revenue.glProfit);
  if (els.stressRevenueMargin) els.stressRevenueMargin.textContent = ratio(scenarioSet.revenue.margin);
  if (els.stressUaPayback) els.stressUaPayback.textContent = formatPaybackMonths(scenarioSet.ua.paybackMonths);
  if (els.stressUaProfit) els.stressUaProfit.textContent = compactMoney(scenarioSet.ua.glProfit);
  if (els.stressUaMargin) els.stressUaMargin.textContent = ratio(scenarioSet.ua.margin);

  stressTestChartInstance = ensureChartInstance(els.stressTestChart, stressTestChartInstance);
  if (!stressTestChartInstance) return;

  const buildCumulativeSeries = (rows) => {
    let running = 0;
    return rows.map((row) => {
      running += Number(row.profit || 0);
      return running;
    });
  };

  stressTestChartInstance.setOption({
    animation: false,
    tooltip: { trigger: "axis", axisPointer: { type: "cross" } },
    legend: { top: 0, icon: "roundRect", textStyle: { color: "#475569" } },
    grid: { left: "4%", right: "4%", top: 36, bottom: 20, containLabel: true },
    xAxis: {
      type: "category",
      boundaryGap: false,
      data: model.monthly.map((row) => formatResultAxisLabel(row.month)),
      axisLabel: { interval: Math.max(Math.floor(model.monthly.length / 8), 0), color: "#94a3b8" },
      axisLine: { lineStyle: { color: "#cbd5e1" } }
    },
    yAxis: {
      type: "value",
      axisLabel: { color: "#94a3b8", formatter: (value) => compactMoney(value) },
      splitLine: { lineStyle: { color: "#e2e8f0", type: "dashed" } }
    },
    series: [
      {
        name: "基准累计项目利润",
        type: "line",
        smooth: true,
        symbol: "none",
        lineStyle: { width: 3, color: "#2563eb" },
        markLine: {
          data: [{ yAxis: 0, name: "回本线" }],
          lineStyle: { color: "#94a3b8", type: "solid", width: 2 },
          label: { formatter: "回本 0 轴", color: "#64748b" },
          symbol: ["none", "none"]
        },
        data: buildCumulativeSeries(scenarioSet.baselineRows)
      },
      {
        name: "收入下移情景",
        type: "line",
        smooth: true,
        symbol: "none",
        lineStyle: { width: 2, type: "dashed", color: "#f59e0b" },
        data: buildCumulativeSeries(scenarioSet.revenueRows)
      },
      {
        name: "UA上浮情景",
        type: "line",
        smooth: true,
        symbol: "none",
        lineStyle: { width: 2, type: "dotted", color: "#ef4444" },
        data: buildCumulativeSeries(scenarioSet.uaRows)
      }
    ]
  }, true);
}

function buildDerivedPostRowsFromSplit() {
  const rowMap = new Map();
  [...state.actualRows.pre, ...state.actualRows.allocated].forEach((row) => {
    const month = normalizeMonth(row.month || row.dt);
    const code = String(row.code || "").trim();
    if (!month || !code) return;
    const account = String(row.account || getTemplateAccountName(code) || "").trim();
    const key = `${month}__${code}`;
    if (!rowMap.has(key)) {
      rowMap.set(key, {
        month,
        code,
        account,
        source: "分摊后",
        amount: 0,
        hc: 0,
        revenue: 0
      });
    }
    const target = rowMap.get(key);
    target.amount += Number(row.amount || 0);
    target.hc = Math.max(Number(target.hc || 0), Number(row.hc || 0));
    target.revenue = Math.max(Number(target.revenue || 0), Number(row.revenue || 0));
  });
  return [...rowMap.values()].sort((a, b) => `${a.month}__${a.code}`.localeCompare(`${b.month}__${b.code}`, "zh-CN"));
}

function getHistoricalPostRows() {
  if (state.actualImportFlags.post && state.actualRows.post.length) {
    return state.actualRows.post;
  }
  if (state.importMode === "split" && (state.actualRows.pre.length || state.actualRows.allocated.length)) {
    return buildDerivedPostRowsFromSplit();
  }
  return state.actualRows.post.length ? state.actualRows.post : buildDerivedPostRowsFromSplit();
}

function buildHistoricalMonthlySeries() {
  const monthMap = new Map();
  const postMetricMap = new Map();
  const directFormalCodes = new Set(["A1.1", "A1.1.1", "A1.2"]);
  const outsourcedCodes = new Set(["A1.4.1", "A1.4.2"]);
  const ensureMonth = (month) => {
    const normalized = normalizeMonth(month);
    if (!normalized) return null;
    if (!monthMap.has(normalized)) {
      monthMap.set(normalized, {
        month: normalized,
        direct: 0,
        allocated: 0,
        merged: 0,
        hc: 0,
        directFormalHc: 0,
        allocatedFormalHc: 0,
        outsourcedHc: 0
      });
    }
    return monthMap.get(normalized);
  };
  const ensurePostMetric = (month) => {
    const normalized = normalizeMonth(month);
    if (!normalized) return null;
    if (!postMetricMap.has(normalized)) {
      postMetricMap.set(normalized, {
        month: normalized,
        revenue: null,
        profit: null,
        hcFromPost: null
      });
    }
    return postMetricMap.get(normalized);
  };

  state.actualRows.pre.forEach((row) => {
    const target = ensureMonth(row.month || row.dt);
    if (!target) return;
    const code = String(row.code || "").trim();
    target.direct += Number(row.amount || 0);
    target.hc = Math.max(target.hc, Number(row.hc || 0));
    if (directFormalCodes.has(code)) {
      target.directFormalHc += Number(row.amount || 0);
    }
    if (outsourcedCodes.has(code)) {
      target.outsourcedHc += Number(row.amount || 0);
    }
  });

  state.actualRows.allocated.forEach((row) => {
    const target = ensureMonth(row.month || row.dt);
    if (!target) return;
    const code = String(row.code || "").trim();
    target.allocated += Number(row.amount || 0);
    target.hc = Math.max(target.hc, Number(row.hc || 0));
    if (directFormalCodes.has(code)) {
      target.allocatedFormalHc += Number(row.amount || 0);
    }
    if (outsourcedCodes.has(code)) {
      target.outsourcedHc += Number(row.amount || 0);
    }
  });

  getHistoricalPostRows().forEach((row) => {
    const target = ensureMonth(row.month || row.dt);
    if (!target) return;
    target.merged += Number(row.amount || 0);
    target.hc = Math.max(target.hc, Number(row.hc || 0));

    const metricTarget = ensurePostMetric(row.month || row.dt);
    if (!metricTarget) return;
    const code = String(row.code || "").trim();
    const amount = Number(row.amount || 0);
    if (code === "A2") {
      metricTarget.revenue = amount;
    } else if (code === "A94") {
      metricTarget.profit = amount;
    } else if (code === "A37" && metricTarget.profit === null) {
      metricTarget.profit = amount;
    } else if (code === "A1") {
      metricTarget.hcFromPost = Math.max(Number(metricTarget.hcFromPost || 0), amount);
    }
  });

  const months = [...monthMap.values()].sort((a, b) => a.month.localeCompare(b.month));
  return months.map((row) => ({
    ...row,
    merged: (() => {
      const metricRow = postMetricMap.get(row.month);
      const profit = Number(metricRow?.profit);
      if (Number.isFinite(profit)) {
        return Math.abs(profit);
      }
      return row.merged || (row.direct + row.allocated);
    })(),
    direct: row.direct || (!state.actualRows.pre.length && !state.actualRows.allocated.length ? (row.merged || 0) : row.direct),
    hc: (() => {
      const metricRow = postMetricMap.get(row.month);
      const hcFromPost = Number(metricRow?.hcFromPost);
      if (Number.isFinite(hcFromPost) && hcFromPost > 0) {
        return hcFromPost;
      }
      return row.hc || 0;
    })(),
    profit: Number(postMetricMap.get(row.month)?.profit ?? 0),
    revenue: Number(postMetricMap.get(row.month)?.revenue ?? 0)
  }));
}

function getHistoricalAvailableMonths() {
  return buildHistoricalMonthlySeries().map((row) => row.month);
}

function syncHistoricalBreakdownRange(months = getHistoricalAvailableMonths()) {
  if (!months.length) {
    state.historicalBreakdownMonth = "";
    state.historicalBreakdownStartMonth = "";
    state.historicalBreakdownEndMonth = "";
    return { startMonth: "", endMonth: "" };
  }
  const fallbackMonth = months.at(-1);
  if (!months.includes(state.historicalBreakdownStartMonth)) {
    state.historicalBreakdownStartMonth = state.historicalBreakdownMonth && months.includes(state.historicalBreakdownMonth)
      ? state.historicalBreakdownMonth
      : fallbackMonth;
  }
  if (!months.includes(state.historicalBreakdownEndMonth)) {
    state.historicalBreakdownEndMonth = state.historicalBreakdownStartMonth || fallbackMonth;
  }
  const startIndex = months.indexOf(state.historicalBreakdownStartMonth);
  const endIndex = months.indexOf(state.historicalBreakdownEndMonth);
  if (startIndex > endIndex) {
    state.historicalBreakdownEndMonth = state.historicalBreakdownStartMonth;
  }
  if (state.historicalBreakdownMode === "single") {
    state.historicalBreakdownStartMonth = state.historicalBreakdownEndMonth || state.historicalBreakdownStartMonth;
  }
  state.historicalBreakdownMonth = state.historicalBreakdownEndMonth;
  return {
    startMonth: state.historicalBreakdownStartMonth,
    endMonth: state.historicalBreakdownEndMonth
  };
}

function getHistoricalBreakdownRows(startMonth, endMonth = startMonth) {
  const normalizedStart = normalizeMonth(startMonth);
  const normalizedEnd = normalizeMonth(endMonth || startMonth);
  if (!normalizedStart || !normalizedEnd) return [];
  const [rangeStart, rangeEnd] = normalizedStart <= normalizedEnd ? [normalizedStart, normalizedEnd] : [normalizedEnd, normalizedStart];
  if (state.actualRows.pre.length || state.actualRows.allocated.length) {
    return [...state.actualRows.pre, ...state.actualRows.allocated]
      .filter((row) => {
        const month = normalizeMonth(row.month || row.dt);
        return month && month >= rangeStart && month <= rangeEnd;
      })
      .map((row) => ({
        label: `${row.account || getTemplateAccountName(row.code)}（${row.source === "分摊值" ? "分摊" : "直属"}）`,
        amount: Number(row.amount || 0),
        source: row.source
      }));
  }
  return getHistoricalPostRows()
    .filter((row) => {
      const month = normalizeMonth(row.month || row.dt);
      return month && month >= rangeStart && month <= rangeEnd;
    })
    .map((row) => ({
      label: row.account || getTemplateAccountName(row.code),
      amount: Number(row.amount || 0),
      source: "分摊后"
    }));
}

function aggregateHistoricalBreakdown(startMonth, endMonth = startMonth, mode = state.historicalBreakdownMode || "cumulative") {
  const groupMap = new Map();
  const normalizedStart = normalizeMonth(startMonth);
  const normalizedEnd = normalizeMonth(endMonth || startMonth);
  const availableMonths = getHistoricalAvailableMonths();
  const rangeMonths = availableMonths.filter((month) => month && month >= normalizedStart && month <= normalizedEnd);
  getHistoricalBreakdownRows(startMonth, endMonth).forEach((row) => {
    if (!groupMap.has(row.label)) {
      groupMap.set(row.label, { label: row.label, amount: 0, source: row.source });
    }
    groupMap.get(row.label).amount += Number(row.amount || 0);
  });
  const divisor = mode === "average" ? Math.max(rangeMonths.length, 1) : 1;
  const rows = [...groupMap.values()]
    .map((row) => ({
      ...row,
      amount: row.amount / divisor
    }))
    .sort((a, b) => b.amount - a.amount);
  if (rows.length <= 8) return rows;
  const head = rows.slice(0, 7);
  const others = rows.slice(7);
  head.push({
    label: "其他",
    amount: sum(others.map((row) => row.amount)),
    source: "其他"
  });
  return head;
}

function buildArcPath(cx, cy, outerRadius, innerRadius, startAngle, endAngle) {
  const polar = (radius, angle) => {
    const radians = ((angle - 90) * Math.PI) / 180;
    return {
      x: cx + radius * Math.cos(radians),
      y: cy + radius * Math.sin(radians)
    };
  };
  const startOuter = polar(outerRadius, startAngle);
  const endOuter = polar(outerRadius, endAngle);
  const startInner = polar(innerRadius, endAngle);
  const endInner = polar(innerRadius, startAngle);
  const largeArc = endAngle - startAngle > 180 ? 1 : 0;
  return [
    `M ${startOuter.x} ${startOuter.y}`,
    `A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${endOuter.x} ${endOuter.y}`,
    `L ${startInner.x} ${startInner.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${endInner.x} ${endInner.y}`,
    "Z"
  ].join(" ");
}

function describeArcStrokePath(cx, cy, radius, startAngle, endAngle) {
  const startRadians = ((startAngle - 90) * Math.PI) / 180;
  const endRadians = ((endAngle - 90) * Math.PI) / 180;
  const start = {
    x: cx + radius * Math.cos(startRadians),
    y: cy + radius * Math.sin(startRadians)
  };
  const end = {
    x: cx + radius * Math.cos(endRadians),
    y: cy + radius * Math.sin(endRadians)
  };
  const largeArc = endAngle - startAngle > 180 ? 1 : 0;
  return `M ${start.x} ${start.y} A ${radius} ${radius} 0 ${largeArc} 1 ${end.x} ${end.y}`;
}

function polarToCartesian(cx, cy, radius, angle) {
  const radians = ((angle - 90) * Math.PI) / 180;
  return {
    x: cx + radius * Math.cos(radians),
    y: cy + radius * Math.sin(radians)
  };
}

function estimateTextWidth(text, fontSize = 12) {
  return String(text || "").length * fontSize * 0.62;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getHistoricalBreakdownChartRows(startMonth, endMonth = startMonth) {
  const normalizedStart = normalizeMonth(startMonth);
  const normalizedEnd = normalizeMonth(endMonth || startMonth);
  if (!normalizedStart || !normalizedEnd) return [];
  const [rangeStart, rangeEnd] = normalizedStart <= normalizedEnd ? [normalizedStart, normalizedEnd] : [normalizedEnd, normalizedStart];
  if (state.actualRows.pre.length || state.actualRows.allocated.length) {
    return [...state.actualRows.pre, ...state.actualRows.allocated]
      .filter((row) => {
        const month = normalizeMonth(row.month || row.dt);
        return month && month >= rangeStart && month <= rangeEnd;
      })
      .map((row) => ({
        code: row.code || "",
        label: row.account || getTemplateAccountName(row.code),
        amount: Number(row.amount || 0),
        source: row.source || ""
      }));
  }
  return getHistoricalPostRows()
    .filter((row) => {
      const month = normalizeMonth(row.month || row.dt);
      return month && month >= rangeStart && month <= rangeEnd;
    })
    .map((row) => ({
      code: row.code || "",
      label: row.account || getTemplateAccountName(row.code),
      amount: Number(row.amount || 0),
      source: "分摊后"
    }));
}

function buildHistoricalBreakdownNestedData(startMonth, endMonth = startMonth, mode = state.historicalBreakdownMode || "cumulative") {
  const availableMonths = getHistoricalAvailableMonths();
  const normalizedStart = normalizeMonth(startMonth);
  const normalizedEnd = normalizeMonth(endMonth || startMonth);
  const rangeMonths = availableMonths.filter((month) => month && month >= normalizedStart && month <= normalizedEnd);
  const divisor = mode === "average" ? Math.max(rangeMonths.length, 1) : 1;
  if (!state.actualRows.pre.length && !state.actualRows.allocated.length) {
    return {
      total: 0,
      innerData: [],
      outerData: [],
      filteredNegative: 0,
      requiresSplitData: true
    };
  }
  const directCodes = {
    labor: new Set(["A14"]),
    outsource: new Set(["A18"]),
    insource: new Set(["A53"]),
    net: new Set(["A5"]),
    profit: new Set(["A94"])
  };
  const allocatedCodes = {
    labor: new Set(["A14"]),
    net: new Set(["A5"]),
    profit: new Set(["A94"])
  };
  const marketingCodes = {
    total: new Set(["A10"]),
    ua: new Set(["A11"]),
    creative: new Set(["A11.1"]),
    testing: new Set(["A11.2"]),
    brand: new Set(["A12"]),
    operations: new Set(["A12.1"])
  };
  const outerConfig = [
    { key: "marketing_ua", parentKey: "marketing", label: "买量投放", color: "#f59e0b" },
    { key: "marketing_creative", parentKey: "marketing", label: "买量素材", color: "#fbbf24" },
    { key: "marketing_testing", parentKey: "marketing", label: "买量测试", color: "#fcd34d" },
    { key: "marketing_brand", parentKey: "marketing", label: "品牌营销", color: "#fde68a" },
    { key: "marketing_operations", parentKey: "marketing", label: "用户运营", color: "#fef3c7" },
    { key: "direct_labor", parentKey: "direct", label: "直属人力成本", color: "#60a5fa" },
    { key: "direct_outsource", parentKey: "direct", label: "直属外包成本", color: "#93c5fd" },
    { key: "direct_insource", parentKey: "direct", label: "直属内包成本", color: "#bfdbfe" },
    { key: "direct_admin", parentKey: "direct", label: "直属行政及其他", color: "#dbeafe" },
    { key: "allocated_labor", parentKey: "allocated", label: "分摊部门人力", color: "#34d399" },
    { key: "allocated_non_labor", parentKey: "allocated", label: "分摊部门非人力", color: "#6ee7b7" }
  ];
  const categoryConfig = [
    { key: "marketing", label: "营销费用", shortLabel: "营销", color: "#f59e0b" },
    { key: "direct", label: "项目组直属", shortLabel: "直属", color: "#3b82f6" },
    { key: "allocated", label: "部门分摊", shortLabel: "分摊", color: "#10b981" }
  ];
  const amountByOuterKey = Object.fromEntries(outerConfig.map((item) => [item.key, 0]));
  const directAmountByCode = new Map();
  const allocatedAmountByCode = new Map();
  const increaseCodeAmount = (targetMap, code, amount) => {
    const normalizedCode = String(code || "").trim().toUpperCase();
    if (!normalizedCode) return;
    targetMap.set(normalizedCode, Number(targetMap.get(normalizedCode) || 0) + amount);
  };
  const getCodeAmount = (targetMap, codes) => {
    const list = Array.isArray(codes) ? codes : [...codes];
    return list.reduce((totalAmount, code) => totalAmount + Number(targetMap.get(String(code || "").trim().toUpperCase()) || 0), 0);
  };
  getHistoricalBreakdownChartRows(startMonth, endMonth).forEach((row) => {
    const amount = Number(row.amount || 0) / divisor;
    const code = String(row.code || "").trim().toUpperCase();
    const targetMap = row.source === "分摊值" ? allocatedAmountByCode : directAmountByCode;
    increaseCodeAmount(targetMap, code, amount);
    if (row.source === "分摊值") {
      if (allocatedCodes.labor.has(code)) amountByOuterKey.allocated_labor += amount;
      return;
    }
    if (marketingCodes.ua.has(code)) amountByOuterKey.marketing_ua += amount;
    else if (marketingCodes.creative.has(code)) amountByOuterKey.marketing_creative += amount;
    else if (marketingCodes.testing.has(code)) amountByOuterKey.marketing_testing += amount;
    else if (marketingCodes.brand.has(code)) amountByOuterKey.marketing_brand += amount;
    else if (marketingCodes.operations.has(code)) amountByOuterKey.marketing_operations += amount;
    if (directCodes.labor.has(code)) amountByOuterKey.direct_labor += amount;
    else if (directCodes.outsource.has(code)) amountByOuterKey.direct_outsource += amount;
    else if (directCodes.insource.has(code)) amountByOuterKey.direct_insource += amount;
  });
  const directNet = getCodeAmount(directAmountByCode, directCodes.net);
  const directProfit = getCodeAmount(directAmountByCode, directCodes.profit);
  const directMarketingTotal = getCodeAmount(directAmountByCode, marketingCodes.total);
  const allocatedNet = getCodeAmount(allocatedAmountByCode, allocatedCodes.net);
  const allocatedProfit = getCodeAmount(allocatedAmountByCode, allocatedCodes.profit);
  amountByOuterKey.direct_admin = directNet - directProfit - directMarketingTotal - amountByOuterKey.direct_labor - amountByOuterKey.direct_outsource - amountByOuterKey.direct_insource;
  amountByOuterKey.allocated_non_labor = allocatedNet - allocatedProfit - amountByOuterKey.allocated_labor;
  const categories = categoryConfig
    .map((category) => {
      const items = outerConfig
        .filter((item) => item.parentKey === category.key)
        .map((item) => ({
          key: item.key,
          label: item.label,
          amount: amountByOuterKey[item.key],
          color: item.color
        }))
        .filter((item) => item.amount !== 0);
      return {
        ...category,
        total: sum(items.map((item) => item.amount)),
        items
      };
    })
    .filter((category) => category.total !== 0);
  const actualTotal = sum(categories.map((item) => item.total));
  return {
    total: actualTotal,
    innerData: categories.map((category) => ({
      name: category.label,
      value: category.total,
      itemStyle: {
        color: category.color,
        borderColor: "#ffffff",
        borderWidth: 3,
        borderRadius: 4,
        shadowBlur: 10,
        shadowColor: "rgba(0,0,0,0.05)"
      },
      rawLabel: category.shortLabel || category.label
    })),
    outerData: categories.flatMap((category) => category.items.map((item) => ({
      name: item.label,
      value: item.amount,
      parentLabel: category.label,
      itemStyle: {
        color: item.color,
        borderColor: "#ffffff",
        borderWidth: 3,
        borderRadius: 6,
        shadowBlur: 15,
        shadowOffsetY: 5,
        shadowColor: "rgba(0,0,0,0.08)"
      },
      parentLabel: category.label
    }))),
    filteredNegative: 0,
    requiresSplitData: false
  };
}

function distributeLabelTargets(entries, minY, maxY, gap) {
  const sorted = [...entries].sort((a, b) => a.targetY - b.targetY);
  let cursor = minY;
  sorted.forEach((entry) => {
    cursor = Math.max(cursor, entry.targetY);
    entry.finalY = Math.min(cursor, maxY);
    cursor = entry.finalY + gap;
  });
  cursor = maxY;
  for (let index = sorted.length - 1; index >= 0; index -= 1) {
    const entry = sorted[index];
    entry.finalY = Math.min(entry.finalY, cursor);
    cursor = entry.finalY - gap;
  }
  return sorted;
}

function bindBreakdownTooltip(panel) {
  if (!panel) return;
  const tooltip = panel.querySelector(".cost-structure-tooltip");
  const canvas = panel.querySelector(".cost-structure-canvas");
  if (!tooltip || !canvas) return;
  const hideTooltip = () => {
    tooltip.classList.remove("is-visible");
  };
  canvas.querySelectorAll("[data-tooltip-title]").forEach((node) => {
    node.addEventListener("mouseenter", () => {
      tooltip.innerHTML = `
        <strong>${escapeHtml(node.dataset.tooltipTitle || "")}</strong>
        <span>${escapeHtml(node.dataset.tooltipValue || "")}</span>
      `;
      tooltip.classList.add("is-visible");
    });
    node.addEventListener("mousemove", (event) => {
      const rect = canvas.getBoundingClientRect();
      tooltip.style.left = `${event.clientX - rect.left + 16}px`;
      tooltip.style.top = `${event.clientY - rect.top - 18}px`;
    });
    node.addEventListener("mouseleave", hideTooltip);
  });
  canvas.addEventListener("mouseleave", hideTooltip);
}

function renderHistoricalBreakdownPanel(monthlySeries) {
  if (!els.historyBreakdownPanel) return;
  const { startMonth, endMonth } = syncHistoricalBreakdownRange(monthlySeries.map((row) => row.month));
  const nestedData = buildHistoricalBreakdownNestedData(startMonth, endMonth, state.historicalBreakdownMode);
  if (nestedData.requiresSplitData) {
    els.historyBreakdownPanel.innerHTML = `<div class="empty-state">该图按“分摊前 + 分摊值”固定口径展示，请先同时导入这两类历史实际损益。</div>`;
    if (historyBreakdownChartInstance) {
      historyBreakdownChartInstance.dispose();
      historyBreakdownChartInstance = null;
    }
    return;
  }
  if (!nestedData.innerData.length || !nestedData.outerData.length) {
    els.historyBreakdownPanel.innerHTML = `<div class="empty-state">请先导入历史实际损益，才能生成费用拆解。</div>`;
    if (historyBreakdownChartInstance) {
      historyBreakdownChartInstance.dispose();
      historyBreakdownChartInstance = null;
    }
    return;
  }

  const titleRange = startMonth === endMonth
    ? formatMonthSlashLabel(startMonth)
    : `${formatMonthSlashLabel(startMonth)} - ${formatMonthSlashLabel(endMonth)}`;
  const modeLabel = state.historicalBreakdownMode === "average"
    ? "月均"
    : state.historicalBreakdownMode === "single"
      ? "单月"
      : "区间累计";
  els.historyBreakdownTitle.textContent = "研发阶段成本结构透视";
  const groupedLegend = nestedData.outerData.reduce((groups, item) => {
    const key = item.parentLabel || "其他";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
    return groups;
  }, new Map());
  const groupedEntries = [...groupedLegend.entries()];
  const innerLookup = new Map(
    nestedData.innerData.map((item) => [item.name, Number(item.value || 0)])
  );
  const categorySummary = groupedEntries.map(([groupName, items]) => {
    const total = Number(innerLookup.get(groupName) || sum(items.map((item) => Number(item.value || 0))));
    return {
      groupName,
      total,
      share: nestedData.total ? total / nestedData.total : 0,
      color: items[0]?.itemStyle?.color || "#cbd5e1"
    };
  });
  const spotlightItems = [...nestedData.outerData]
    .sort((a, b) => Math.abs(Number(b.value || 0)) - Math.abs(Number(a.value || 0)))
    .slice(0, 5);
  els.historyBreakdownPanel.innerHTML = `
    <div class="cost-structure-shell">
      <div class="cost-structure-meta">
        <span class="cost-structure-kicker">${escapeHtml(`${titleRange} · ${modeLabel}`)}</span>
        <span class="cost-structure-total">拆分口径总额 ${escapeHtml(formatWan(nestedData.total))}</span>
        <span class="cost-structure-note">口径：仅按设定取数口径读取实际数，不对齐 A94，不做归一</span>
      </div>
      <div class="cost-structure-summary-grid">
        ${categorySummary.map((item) => `
          <article class="cost-structure-summary-card">
            <span class="cost-structure-summary-dot" style="background:${item.color}"></span>
            <div class="cost-structure-summary-copy">
              <strong>${escapeHtml(item.groupName)}</strong>
              <span>${escapeHtml(formatWan(item.total))}</span>
            </div>
            <b>${escapeHtml(ratio(item.share))}</b>
          </article>
        `).join("")}
      </div>
      <div class="cost-structure-canvas">
        <div class="cost-structure-layout">
          <div class="cost-structure-echart-wrap">
            <div class="cost-structure-spotlight">
              <div class="cost-structure-spotlight-head">
                <span>重点科目</span>
                <b>按绝对值排序</b>
              </div>
              <div class="cost-structure-spotlight-list">
                ${spotlightItems.map((item) => `
                  <div class="cost-structure-spotlight-item">
                    <span class="cost-structure-spotlight-dot" style="background:${item.itemStyle?.color || "#cbd5e1"}"></span>
                    <strong>${escapeHtml(item.name)}</strong>
                    <span>${escapeHtml(formatWan(item.value, 1))}</span>
                  </div>
                `).join("")}
              </div>
            </div>
            <div class="cost-structure-echart-stage">
              <div class="cost-structure-echart" id="history-breakdown-echart" aria-label="研发阶段成本结构透视图"></div>
              <div class="cost-structure-center-overlay" aria-hidden="true">
                <span>${escapeHtml(formatWan(nestedData.total, 1))}</span>
              </div>
            </div>
          </div>
          <aside class="cost-structure-sidepanel" aria-label="成本结构明细说明">
            <div class="cost-structure-sidepanel-head">
              <span>结构明细</span>
              <b>单位：USD</b>
            </div>
            ${groupedEntries.map(([groupName, items]) => {
              const groupTotal = Number(innerLookup.get(groupName) || sum(items.map((item) => Number(item.value || 0))));
              const groupShare = nestedData.total ? groupTotal / nestedData.total : 0;
              return `
              <section class="cost-structure-group">
                <div class="cost-structure-group-head">
                  <div>
                    <h4 class="cost-structure-group-title">${escapeHtml(groupName)}</h4>
                    <p class="cost-structure-group-subtitle">${escapeHtml(formatWan(groupTotal))}</p>
                  </div>
                  <span class="cost-structure-group-badge">${escapeHtml(ratio(groupShare))}</span>
                </div>
                <div class="cost-structure-group-list">
                  ${items.map((item) => `
                    <div class="cost-structure-legend-item">
                      <span class="cost-structure-legend-dot" style="background:${item.itemStyle?.color || "#cbd5e1"}"></span>
                      <div class="cost-structure-legend-copy">
                        <strong>${escapeHtml(item.name)}</strong>
                        <span>${escapeHtml(formatWan(item.value))}</span>
                      </div>
                      <b>${escapeHtml(ratio(nestedData.total ? Number(item.value || 0) / nestedData.total : 0))}</b>
                    </div>
                  `).join("")}
                </div>
              </section>
            `;
            }).join("")}
          </aside>
        </div>
      </div>
    </div>
  `;
  const chartHost = document.getElementById("history-breakdown-echart");
  if (!chartHost) return;
  if (!window.echarts) {
    chartHost.innerHTML = `<div class="empty-state">ECharts 资源加载失败，请刷新页面重试。</div>`;
    return;
  }
  if (historyBreakdownChartInstance) {
    historyBreakdownChartInstance.dispose();
  }
  const formatRichLabel = (params) => {
    const value = Number(params?.value || 0);
    const percent = nestedData.total ? value / nestedData.total : 0;
    const side = params?.data?.labelSide === "left" ? "left" : "right";
    const nameKey = side === "left" ? "nameLeft" : "nameRight";
    const amountKey = side === "left" ? "amountLeft" : "amountRight";
    const badgeKey = side === "left" ? "badgeLeft" : "badgeRight";
    return `{${nameKey}|${params?.name || ""}}\n{${amountKey}|${formatWan(value, 1)}}\n{${badgeKey}|${ratio(percent)}}`;
  };
  const outerTotal = sum(nestedData.outerData.map((item) => Number(item.value || 0)));
  let outerCumulative = 0;
  const outerLabelSideMap = new Map();
  nestedData.outerData.forEach((item) => {
    const itemValue = Number(item.value || 0);
    const midRatio = outerTotal ? (outerCumulative + itemValue / 2) / outerTotal : 0;
    const angle = (105 - midRatio * 360) * Math.PI / 180;
    const side = Math.cos(angle) < 0 ? "left" : "right";
    outerLabelSideMap.set(item.name, side);
    outerCumulative += itemValue;
  });
  const outerLabelNames = new Set(
    [...nestedData.outerData]
      .sort((a, b) => Math.abs(Number(b.value || 0)) - Math.abs(Number(a.value || 0)))
      .slice(0, 4)
      .map((item) => item.name)
  );
  try {
    historyBreakdownChartInstance = window.echarts.init(chartHost, null, { renderer: "canvas" });
    historyBreakdownChartInstance.setOption({
    animationDuration: 500,
    animationEasing: "cubicOut",
    animationDurationUpdate: 240,
    graphic: [],
    tooltip: {
      trigger: "item",
      backgroundColor: "rgba(255, 255, 255, 0.95)",
      borderColor: "#e2e8f0",
      borderWidth: 1,
      textStyle: {
        color: "#1e293b",
        fontSize: 12
      },
      extraCssText: "box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border-radius: 8px; backdrop-filter: blur(10px);",
      formatter: (params) => {
        const percent = nestedData.total ? Number(params.value || 0) / nestedData.total : 0;
        const parentText = params.data?.parentLabel ? `<div style=\"margin-bottom:4px;color:#64748b;\">${escapeHtml(params.data.parentLabel)}</div>` : "";
        return `${parentText}<div style="font-weight:700;margin-bottom:4px;">${escapeHtml(params.name || "")}</div><div>${escapeHtml(formatWan(params.value))} · ${escapeHtml(ratio(percent))}</div>`;
      }
    },
    series: [
      {
        type: "pie",
        radius: ["20%", "43%"],
        center: ["43.5%", "57%"],
        avoidLabelOverlap: false,
        selectedMode: false,
        minAngle: 6,
        startAngle: 105,
        z: 2,
        cursor: "pointer",
        label: {
          show: true,
          position: "inside",
          formatter: (params) => {
            const percent = nestedData.total ? Number(params.value || 0) / nestedData.total : 0;
            if (percent <= 0.07) {
              return `{innerNameSmall|${params.data.rawLabel}}\n{innerPctSmall|${ratio(percent)}}`;
            }
            return `{innerName|${params.data.rawLabel}}\n{innerPct|${ratio(percent)}}`;
          },
          rich: {
            innerName: {
              color: "#dbeafe",
              fontSize: 11,
              fontWeight: 700,
              lineHeight: 15,
              align: "center"
            },
            innerPct: {
              color: "#ffffff",
              fontSize: 16,
              fontWeight: 700,
              lineHeight: 20,
              align: "center"
            },
            innerNameSmall: {
              color: "#dbeafe",
              fontSize: 9,
              fontWeight: 700,
              lineHeight: 11,
              align: "center"
            },
            innerPctSmall: {
              color: "#ffffff",
              fontSize: 12,
              fontWeight: 700,
              lineHeight: 14,
              align: "center"
            }
          }
        },
        labelLine: { show: false },
        emphasis: {
          scale: true,
          scaleSize: 6,
          focus: "self",
          itemStyle: {
            shadowBlur: 18,
            shadowColor: "rgba(15,23,42,0.14)"
          },
          label: {
            rich: {
              innerName: {
                fontSize: 11
              },
              innerPct: {
                fontSize: 17
              },
              innerNameSmall: {
                fontSize: 9
              },
              innerPctSmall: {
                fontSize: 12
              }
            }
          }
        },
        data: nestedData.innerData
      },
      {
        type: "pie",
        radius: ["50%", "82%"],
        center: ["43.5%", "57%"],
        avoidLabelOverlap: true,
        minAngle: 3,
        startAngle: 105,
        z: 3,
        cursor: "pointer",
        label: {
          show: true,
          formatter: formatRichLabel,
          alignTo: "labelLine",
          edgeDistance: 6,
          bleedMargin: 4,
          width: 116,
          overflow: "truncate",
          lineHeight: 16,
          rich: {
            nameLeft: {
              color: "#334155",
              fontSize: 12,
              fontWeight: 700,
              width: 116,
              lineHeight: 16,
              align: "right",
              padding: [0, 0, 3, 0]
            },
            nameRight: {
              color: "#334155",
              fontSize: 12,
              fontWeight: 700,
              width: 116,
              lineHeight: 16,
              align: "left",
              padding: [0, 0, 3, 0]
            },
            amountLeft: {
              color: "#64748b",
              fontSize: 11,
              fontWeight: 600,
              width: 116,
              lineHeight: 15,
              align: "right",
              padding: [0, 0, 3, 0]
            },
            amountRight: {
              color: "#64748b",
              fontSize: 11,
              fontWeight: 600,
              width: 116,
              lineHeight: 15,
              align: "left",
              padding: [0, 0, 3, 0]
            },
            badgeLeft: {
              color: "#cbd5e1",
              fontSize: 11,
              fontWeight: 700,
              backgroundColor: "#1e293b",
              borderRadius: 4,
              align: "right",
              padding: [2, 5]
            },
            badgeRight: {
              color: "#cbd5e1",
              fontSize: 11,
              fontWeight: 700,
              backgroundColor: "#1e293b",
              borderRadius: 4,
              align: "left",
              padding: [2, 5]
            }
          }
        },
        labelLayout: {
          hideOverlap: true,
          moveOverlap: "shiftY"
        },
        labelLine: {
          show: true,
          length: 12,
          length2: 10,
          smooth: 0.2,
          lineStyle: {
            color: "#cbd5e1",
            width: 1.2
          }
        },
        emphasis: {
          scale: true,
          scaleSize: 10,
          focus: "self",
          itemStyle: {
            shadowBlur: 22,
            shadowOffsetY: 8,
            shadowColor: "rgba(15,23,42,0.18)"
          },
          labelLine: {
            lineStyle: {
              width: 2,
              color: "#94a3b8"
            }
          }
        },
        data: nestedData.outerData.map((item) => ({
          ...item,
          labelSide: outerLabelSideMap.get(item.name) || "right",
          label: {
            show: outerLabelNames.has(item.name),
            align: (outerLabelSideMap.get(item.name) || "right") === "left" ? "right" : "left"
          },
          labelLine: {
            show: outerLabelNames.has(item.name)
          }
        }))
      }
    ]
    });
    requestAnimationFrame(() => historyBreakdownChartInstance?.resize());
  } catch (error) {
    console.error("Failed to render historical breakdown chart", error);
    chartHost.innerHTML = `<div class="empty-state">双环图渲染失败：${escapeHtml(error?.message || "未知错误")}</div>`;
    if (historyBreakdownChartInstance) {
      historyBreakdownChartInstance.dispose();
      historyBreakdownChartInstance = null;
    }
  }
}

function renderHistoricalTrendChart(monthlySeries) {
  const historyTrendChart = document.getElementById("history-trend-chart");
  if (!historyTrendChart) return;
  const metricMonthlyMap = getHistoricalMetricMonthlyMap("A94");
  const rows = monthlySeries.slice(-12).map((row) => {
    const mappedSpend = Number(metricMonthlyMap.get(row.month));
    const spend = Number.isFinite(mappedSpend) ? Math.abs(mappedSpend) : Math.abs(Number(row.merged || 0));
    return {
      ...row,
      spend
    };
  });
  if (!rows.length) {
    historyTrendChart.innerHTML = "";
    return;
  }
  const width = 420;
  const height = 150;
  const chartLeft = 28;
  const chartRight = 16;
  const chartTop = 10;
  const chartBottom = 8;
  const baseY = height - chartBottom;
  const chartWidth = width - chartLeft - chartRight;
  const usableHeight = 126;
  const maxSpend = Math.max(...rows.map((row) => Math.abs(Number(row.spend || 0))), 1);
  const maxHc = Math.max(...rows.map((row) => row.hc), 1);
  const step = chartWidth / Math.max(rows.length, 1);
  const barWidth = Math.min(step * 0.42, 20);
  const pointX = (index) => chartLeft + step * index + step / 2;
  const barX = (index) => pointX(index) - barWidth / 2;
  const spendY = (value) => baseY - (Math.abs(Number(value || 0)) / maxSpend) * usableHeight;
  const hcY = (value) => baseY - (value / maxHc) * usableHeight;
  const linePath = rows.map((row, index) => `${index === 0 ? "M" : "L"} ${pointX(index)} ${hcY(row.hc)}`).join(" ");
  historyTrendChart.innerHTML = `
    <line x1="${chartLeft}" y1="${baseY}" x2="${width - chartRight}" y2="${baseY}" stroke="rgba(19,34,56,0.18)"></line>
    <line x1="${chartLeft}" y1="${chartTop + 12}" x2="${width - chartRight}" y2="${chartTop + 12}" stroke="rgba(19,34,56,0.08)" stroke-dasharray="4 4"></line>
    ${rows.map((row, index) => `
      <rect class="hero-trend-bar" data-trend-index="${index}" x="${barX(index)}" y="${spendY(row.spend)}" width="${barWidth}" height="${Math.max(baseY - spendY(row.spend), row.spend > 0 ? 1.5 : 0)}" rx="5" fill="#4a98ea" opacity="0.92"></rect>
    `).join("")}
    <path d="${linePath}" fill="none" stroke="#ef6c6c" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"></path>
    ${rows.map((row, index) => `<circle class="hero-trend-node" data-trend-index="${index}" cx="${pointX(index)}" cy="${hcY(row.hc)}" r="${index === rows.length - 1 ? 3.8 : 3}" fill="#ffffff" stroke="#ef6c6c" stroke-width="${index === rows.length - 1 ? 2 : 1.6}"></circle>`).join("")}
  `;

  const svgNs = "http://www.w3.org/2000/svg";
  rows.forEach((row, index) => {
    const bar = historyTrendChart.querySelector(`.hero-trend-bar[data-trend-index="${index}"]`);
    if (bar) {
      const title = document.createElementNS(svgNs, "title");
      title.textContent = `${formatMonthSlashLabel(row.month)} 支出：${formatWan(row.spend, 1)}`;
      bar.appendChild(title);
    }
    const node = historyTrendChart.querySelector(`.hero-trend-node[data-trend-index="${index}"]`);
    if (node) {
      const title = document.createElementNS(svgNs, "title");
      title.textContent = `${formatMonthSlashLabel(row.month)} HC：${compactDecimal(row.hc, 1)} 人`;
      node.appendChild(title);
    }
  });
}

function getHistoricalOverviewMetrics(monthlySeries) {
  const rows = monthlySeries.filter((row) => row.month);
  const latestRow = rows.at(-1) || { merged: 0, hc: 0 };
  const recentRows = rows.slice(-Math.max(Number(state.historicalBurnWindow) || 3, 1));
  const netRevenueMonthlyMap = getHistoricalMetricMonthlyMap("A5");
  const metricMonthlyMap = getHistoricalMetricMonthlyMap("A94");
  const marketingTotalMonthlyMap = getHistoricalMetricMonthlyMap("A10");
  const marketingDetailMonthlyMap = getHistoricalMetricMonthlyMap(["A11", "A11.1", "A11.2", "A12", "A12.1"]);
  const nonMarketingMonthlyMap = getHistoricalMetricMonthlyMap(["A7", "A98", "A13", "A53", "A35", "A45"]);
  const metricSeries = rows.map((row) => ({
    month: row.month,
    value: Number(metricMonthlyMap.get(row.month) ?? row.profit ?? 0)
  }));
  const marketingSeries = rows.map((row) => ({
    month: row.month,
    value: Number(marketingTotalMonthlyMap.has(row.month) ? marketingTotalMonthlyMap.get(row.month) : (marketingDetailMonthlyMap.get(row.month) ?? 0))
  }));
  const rdCostSeries = rows.map((row) => {
    const netRevenue = Number(netRevenueMonthlyMap.get(row.month) ?? row.revenue ?? 0);
    const profit = Number(metricMonthlyMap.get(row.month) ?? row.profit ?? 0);
    const marketingCost = Number(marketingTotalMonthlyMap.has(row.month) ? marketingTotalMonthlyMap.get(row.month) : (marketingDetailMonthlyMap.get(row.month) ?? 0));
    return {
      month: row.month,
      value: netRevenue - profit - marketingCost
    };
  });
  const nonMarketingSeries = rows.map((row) => ({
    month: row.month,
    value: Math.max(Number(nonMarketingMonthlyMap.get(row.month) ?? 0), 0)
  }));
  const recentMetricRows = metricSeries.slice(-Math.max(Number(state.historicalBurnWindow) || 3, 1));
  const recentMarketingRows = marketingSeries.slice(-Math.max(Number(state.historicalBurnWindow) || 3, 1));
  const recentRdCostRows = rdCostSeries.slice(-Math.max(Number(state.historicalBurnWindow) || 3, 1));
  const recentNonMarketingRows = nonMarketingSeries.slice(-Math.max(Number(state.historicalBurnWindow) || 3, 1));
  const burnRate = avg(recentMetricRows.map((row) => Math.abs(row.value)));
  const avgRecentHc = avg(recentRows.map((row) => row.hc).filter((value) => value > 0));
  const latestMonthHc = Number(latestRow.hc || 0);
  const avgDirectFormalHc = avg(recentRows.map((row) => row.directFormalHc).filter((value) => value > 0));
  const avgAllocatedFormalHc = avg(recentRows.map((row) => row.allocatedFormalHc).filter((value) => value > 0));
  const avgOutsourcedHc = avg(recentRows.map((row) => row.outsourcedHc).filter((value) => value > 0));
  const cumulativeProfit = sum(metricSeries.map((row) => Number(row.value || 0)));
  const cumulativeSpend = Math.abs(cumulativeProfit);
  const avgRecentMarketingCost = avg(recentMarketingRows.map((row) => row.value));
  const avgRecentRdCost = avg(recentRdCostRows.map((row) => Number(row.value || 0)));
  const avgRecentNonMarketingCost = avg(recentNonMarketingRows.map((row) => row.value));
  const stackedCostBase = avgRecentMarketingCost + avgRecentNonMarketingCost;
  const marketingShare = stackedCostBase > 0 ? clamp(avgRecentMarketingCost / stackedCostBase, 0, 1) : 0;
  const nonMarketingShare = stackedCostBase > 0 ? clamp(avgRecentNonMarketingCost / stackedCostBase, 0, 1) : 1;
  const hcMarginalCost = avgRecentHc ? burnRate / avgRecentHc : 0;
  const hcMarginalRdCost = avgRecentHc ? avgRecentRdCost / avgRecentHc : 0;
  const hcMarginalMarketingCost = avgRecentHc ? avgRecentMarketingCost / avgRecentHc : 0;
  const sparklineValues = metricSeries.slice(-6).map((row) => Math.abs(row.value));
  const latestMetric = Math.abs(Number(metricSeries.at(-1)?.value || 0));
  const previousMetric = Math.abs(Number(metricSeries.at(-2)?.value || 0));
  const burnRateMoM = previousMetric ? (latestMetric - previousMetric) / previousMetric : 0;
  return {
    latestMonth: latestRow.month || "",
    cumulativeSpend,
    burnRate,
    avgRecentHc,
    latestMonthHc,
    avgDirectFormalHc,
    avgAllocatedFormalHc,
    avgOutsourcedHc,
    hcMarginalCost,
    avgRecentRdCost,
    hcMarginalRdCost,
    hcMarginalMarketingCost,
    avgRecentMarketingCost,
    avgRecentNonMarketingCost,
    marketingShare,
    nonMarketingShare,
    sparklineValues,
    burnRateMoM,
    monthlyRows: rows
  };
}

function getBudgetAlertMetrics(overview) {
  const latestHistoricalMonth = normalizeMonth(overview.latestMonth || getHistoricalAvailableMonths().at(-1) || "");
  if (!state.budgetAlertLaunchMonth && latestHistoricalMonth) {
    state.budgetAlertLaunchMonth = addMonths(latestHistoricalMonth, 6);
  }
  const launchMonthDate = monthKeyToDate(state.budgetAlertLaunchMonth);
  const latestMonthDate = monthKeyToDate(latestHistoricalMonth);
  const monthsToLaunch = Math.max(Number(monthDiff(latestMonthDate, launchMonthDate)) || 0, 0);
  const projectedMonthlyCost = overview.latestMonthHc * (1 + state.budgetAlertHcGrowthRate) * overview.hcMarginalRdCost;
  const projectedToLaunchCost = monthsToLaunch * projectedMonthlyCost + state.budgetAlertMarketingBudget;
  const cumulativeProjectedCost = overview.cumulativeSpend + projectedToLaunchCost;
  const remainingBudget = state.projectBudgetAmount - overview.cumulativeSpend;
  const budgetGap = Math.max(cumulativeProjectedCost - state.projectBudgetAmount, 0);
  const remainingBudgetGap = remainingBudget - projectedToLaunchCost;
  const budgetProgressRaw = state.projectBudgetAmount > 0 ? overview.cumulativeSpend / state.projectBudgetAmount : 0;
  const budgetProgress = clamp(budgetProgressRaw, 0, 1);
  let statusTone = "neutral";
  let statusLabel = "待补充预算";
  let statusText = "请先填写项目预算（USD），系统会自动判断预算是否足够支撑到上线。";

  if (state.projectBudgetAmount > 0 && state.budgetAlertLaunchMonth) {
    if (budgetGap > 0) {
      statusTone = "negative";
      statusLabel = "预计超预算";
      statusText = `按当前计划，预计需补充申请预算 ${formatWan(budgetGap, 1)}。`;
    } else if (cumulativeProjectedCost >= state.projectBudgetAmount * 0.85) {
      statusTone = "warning";
      statusLabel = "接近预算上限";
      statusText = `按当前计划，预算仍可覆盖上线前投入，但已接近预算上限。`;
    } else {
      statusTone = "positive";
      statusLabel = "预算充足";
      statusText = "按当前计划，当前项目预算可覆盖到上线阶段预计累计费用。";
    }
  } else if (state.projectBudgetAmount > 0) {
    statusText = "请先填写预计上线时间，再生成预算预警结论。";
  }

  return {
    latestHistoricalMonth,
    monthsToLaunch,
    projectedMonthlyCost,
    projectedToLaunchCost,
    cumulativeProjectedCost,
    budgetGap,
    budgetProgress,
    budgetProgressRaw,
    remainingBudget,
    remainingBudgetGap,
    statusTone,
    statusLabel,
    statusText
  };
}

function renderHistoricalOverview() {
  const monthlySeries = buildHistoricalMonthlySeries();
  const availableMonths = monthlySeries.map((row) => row.month);
  const { startMonth, endMonth } = syncHistoricalBreakdownRange(availableMonths);

  if (els.historyBreakdownStartMonth) {
    els.historyBreakdownStartMonth.innerHTML = availableMonths.length
      ? availableMonths.map((month) => `<option value="${month}">${formatMonthSlashLabel(month)}</option>`).join("")
      : `<option value="">暂无历史月份</option>`;
    els.historyBreakdownStartMonth.value = startMonth || "";
    els.historyBreakdownStartMonth.disabled = state.historicalBreakdownMode === "single";
  }
  if (els.historyBreakdownEndMonth) {
    els.historyBreakdownEndMonth.innerHTML = availableMonths.length
      ? availableMonths.map((month) => `<option value="${month}">${formatMonthSlashLabel(month)}</option>`).join("")
      : `<option value="">暂无历史月份</option>`;
    els.historyBreakdownEndMonth.value = endMonth || "";
  }
  if (els.historyBreakdownModeSwitch) {
    els.historyBreakdownModeSwitch.querySelectorAll("[data-breakdown-mode]").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.breakdownMode === state.historicalBreakdownMode);
    });
  }
  if (els.historyBurnWindow) {
    els.historyBurnWindow.value = String(state.historicalBurnWindow);
  }
  if (els.historyProjectBudget) {
    els.historyProjectBudget.value = state.projectBudgetAmount > 0 ? integerWithThousands(state.projectBudgetAmount) : "";
  }

  const overview = getHistoricalOverviewMetrics(monthlySeries);
  if (!monthlySeries.length) {
    if (els.historyOverviewSummary) els.historyOverviewSummary.innerHTML = `<div class="empty-state">请先导入历史实际损益，才能生成概览与趋势。</div>`;
    if (els.historyBreakdownPanel) els.historyBreakdownPanel.innerHTML = `<div class="empty-state">暂无历史费用拆解数据。</div>`;
    const historyTrendChart = document.getElementById("history-trend-chart");
    if (historyTrendChart) historyTrendChart.innerHTML = "";
    return overview;
  }

  if (els.historyOverviewSummary) {
    const spendParts = metricParts(overview.cumulativeSpend, "万", 1, 10000);
    const burnParts = metricParts(overview.burnRate, "万/月", 1, 10000);
    const hcParts = metricParts(overview.avgRecentHc, "人", 1, 1);
    const hcCostParts = metricParts(overview.hcMarginalCost, "万/人/月", 2, 10000);
    const hcRdCostParts = metricParts(overview.hcMarginalRdCost, "万/人/月", 2, 10000);
    const hcMarketingCostParts = metricParts(overview.hcMarginalMarketingCost, "万/人/月", 2, 10000);
    const directHcParts = metricParts(overview.avgDirectFormalHc, "人", 1, 1);
    const allocatedHcParts = metricParts(overview.avgAllocatedFormalHc, "人", 1, 1);
    const outsourcedHcParts = metricParts(overview.avgOutsourcedHc, "人", 1, 1);
    const budgetProgressRaw = state.projectBudgetAmount > 0 ? overview.cumulativeSpend / state.projectBudgetAmount : 0;
    const budgetAlert = getBudgetAlertMetrics(overview);
    const sparkline = buildSparklineGeometry(overview.sparklineValues);
    const budgetInputValue = state.projectBudgetAmount > 0 ? formatWan(state.projectBudgetAmount, 1) : "待填写";
    const remainingBudgetValue = state.projectBudgetAmount > 0 ? formatSignedWan(budgetAlert.remainingBudget, 1) : "待填写";
    const consumedBudgetValue = state.projectBudgetAmount > 0 ? `${Math.round(budgetProgressRaw * 100)}%` : "待填写";
    const summaryStatusValue = state.projectBudgetAmount > 0 && state.budgetAlertLaunchMonth
      ? budgetAlert.statusTone === "negative"
        ? `${budgetAlert.statusLabel} ${formatWan(budgetAlert.budgetGap, 1)}`
        : budgetAlert.statusLabel
      : budgetAlert.statusLabel;
    const trendRows = monthlySeries.slice(-12);
    const trendLabelMonths = trendRows.length
      ? [trendRows[0], trendRows[Math.max(Math.floor((trendRows.length - 1) / 2), 0)], trendRows[trendRows.length - 1]]
          .map((row) => formatMonthSlashLabel(row.month))
      : [];
    els.historyOverviewSummary.innerHTML = `
      <article class="step1-summary-card hero-metric-card hero-budget-card">
        <div class="hero-card-top">
          <h3 class="hero-card-title">累计研发花费与预算预警</h3>
          <p class="hero-card-subtitle">历史累计消耗与上线前预算压力。</p>
          <div class="hero-budget-kpi-row">
            <div class="hero-budget-kpi-main">
              <div class="hero-card-value">
                <strong class="hero-card-number">${spendParts.value}</strong>
                <span class="hero-card-unit">${spendParts.unit}</span>
              </div>
              <span class="hero-card-caption">累计研发花费总额</span>
            </div>
            <div class="hero-budget-kpi-side">
              <span><em>已消耗预算</em><b>${consumedBudgetValue}</b></span>
              <span><em>剩余预算</em><b>${remainingBudgetValue}</b></span>
            </div>
          </div>
        </div>
        <div class="hero-card-bottom hero-budget-bottom">
          <div class="hero-progress-track hero-progress-track-lg">
            <div class="hero-progress-fill hero-progress-fill-blue" style="width:${budgetAlert.budgetProgress * 100}%"></div>
          </div>
          <div class="hero-budget-meta">
            <span><em>项目预算</em><b>${budgetInputValue}</b></span>
            <span><em>${state.projectBudgetAmount > 0 ? "预算进度" : "提示"}</em><b>${state.projectBudgetAmount > 0 ? `已消耗预算 ${consumedBudgetValue}` : "请先填写项目预算"}</b></span>
          </div>
          <section class="hero-budget-summary hero-budget-summary-${budgetAlert.statusTone}">
            <div class="hero-budget-summary-head">
              <strong>预算预警摘要</strong>
              <span class="hero-budget-status-pill hero-budget-status-pill-${budgetAlert.statusTone}">${summaryStatusValue}</span>
            </div>
            <div class="hero-budget-summary-grid">
              <span><em>预计上线时间</em><b>${state.budgetAlertLaunchMonth ? formatMonthSlashLabel(state.budgetAlertLaunchMonth) : "待填写"}</b></span>
              <span><em>到上线阶段预计累计费用</em><b>${state.projectBudgetAmount > 0 && state.budgetAlertLaunchMonth ? formatWan(budgetAlert.cumulativeProjectedCost, 1) : "待计算"}</b></span>
              <span class="hero-budget-summary-wide"><em>状态说明</em><b>${budgetAlert.statusText}</b></span>
            </div>
          </section>
          <div class="hero-budget-expand is-open">
            <div class="hero-budget-input-grid">
              <label class="field">
                <span>预计上线时间</span>
                <input id="history-budget-alert-launch-month" type="month" value="${escapeHtml(state.budgetAlertLaunchMonth || "")}" />
              </label>
              <label class="field">
                <span>上线前营销预算（USD）</span>
                <input id="history-budget-alert-marketing-budget" type="text" inputmode="numeric" placeholder="请输入上线前营销预算" value="${state.budgetAlertMarketingBudget > 0 ? integerWithThousands(state.budgetAlertMarketingBudget) : ""}" />
              </label>
              <label class="field hero-budget-range-field">
                <span>上线前 HC 扩张率</span>
                <div class="hero-budget-range">
                  <input id="history-budget-alert-hc-growth-range" type="range" min="0" max="100" step="5" value="${Math.round(state.budgetAlertHcGrowthRate * 100)}" />
                  <strong>${state.budgetAlertHcGrowthRate > 0 ? "+" : ""}${Math.round(state.budgetAlertHcGrowthRate * 100)}%</strong>
                </div>
              </label>
            </div>
            <div class="hero-budget-result-grid">
              <article class="hero-budget-result">
                <span>月预估研发费用</span>
                <strong>${formatWanPerMonth(budgetAlert.projectedMonthlyCost, 1)}</strong>
              </article>
              <article class="hero-budget-result">
                <span>到上线阶段预估费用需求</span>
                <strong>${formatWan(budgetAlert.projectedToLaunchCost, 1)}</strong>
              </article>
              <article class="hero-budget-result${budgetAlert.remainingBudgetGap < 0 ? " is-danger" : ""}">
                <span>需补充申请预算</span>
                <strong>${formatSignedWan(budgetAlert.remainingBudgetGap, 1)}</strong>
              </article>
            </div>
          </div>
        </div>
      </article>
      <div class="step1-summary-side-stack">
        <article class="step1-summary-card hero-metric-card hero-side-card hero-side-card-burn">
          <div class="hero-card-top">
            <h3 class="hero-card-title">近期月均花费</h3>
          </div>
          <div class="hero-card-middle">
            <div class="hero-side-value-row">
              <div class="hero-card-value">
                <strong class="hero-card-number">${burnParts.value}</strong>
                <span class="hero-card-unit">${burnParts.unit}</span>
              </div>
              <span class="hero-pill hero-pill-red">环比 ${formatSignedPercent(overview.burnRateMoM, 1)}</span>
            </div>
            <span class="hero-card-caption">基于最近${state.historicalBurnWindow}个月平均</span>
          </div>
        </article>
        <article class="step1-summary-card hero-metric-card hero-side-card hero-side-card-hc">
          <div class="hero-card-top">
            <h3 class="hero-card-title">近期月均 HC 规模</h3>
          </div>
          <div class="hero-card-middle">
            <div class="hero-card-value">
              <strong class="hero-card-number">${hcParts.value}</strong>
              <span class="hero-card-unit">人/月</span>
            </div>
            <span class="hero-card-caption">直属、分摊、外包</span>
          </div>
          <div class="hero-card-bottom hero-side-bottom">
            <div class="hero-tag-row hero-tag-row-compact">
              <span class="hero-tag hero-tag-blue">
                <i class="hero-tag-dot"></i>
                <em>直属</em>
                <b>${directHcParts.value}</b>
              </span>
              <span class="hero-tag hero-tag-emerald">
                <i class="hero-tag-dot"></i>
                <em>分摊</em>
                <b>${allocatedHcParts.value}</b>
              </span>
              <span class="hero-tag hero-tag-amber">
                <i class="hero-tag-dot"></i>
                <em>外包</em>
                <b>${outsourcedHcParts.value}</b>
              </span>
            </div>
            <div class="hero-cost-split">
              <div class="hero-cost-split-head">
                <span>HC 综合边际月成本拆分</span>
              </div>
              <div class="hero-stacked-track hero-cost-split-track" aria-hidden="true">
                <div class="hero-cost-split-fill hero-cost-split-fill-rd" style="width:${Math.max(0, Math.min(100, overview.hcMarginalCost > 0 ? overview.hcMarginalRdCost / overview.hcMarginalCost * 100 : 0))}%"></div>
                <div class="hero-cost-split-fill hero-cost-split-fill-marketing" style="width:${Math.max(0, Math.min(100, overview.hcMarginalCost > 0 ? overview.hcMarginalMarketingCost / overview.hcMarginalCost * 100 : 0))}%"></div>
              </div>
              <div class="hero-cost-split-legend">
                <span class="hero-cost-split-item hero-cost-split-item-rd">
                  <i class="hero-cost-split-dot"></i>
                  <em>研发</em>
                  <b>${hcRdCostParts.value}${hcRdCostParts.unit}</b>
                </span>
                <span class="hero-cost-split-item hero-cost-split-item-marketing">
                  <i class="hero-cost-split-dot"></i>
                  <em>营销</em>
                  <b>${hcMarketingCostParts.value}${hcMarketingCostParts.unit}</b>
                </span>
              </div>
            </div>
          </div>
        </article>
        <article class="step1-summary-card hero-metric-card hero-side-card hero-trend-card">
          <div class="hero-card-top">
            <h3 class="hero-card-title">HC 与成本趋势</h3>
          </div>
          <div class="hero-card-middle hero-trend-middle">
            <div class="svg-wrap hero-trend-wrap">
              <svg id="history-trend-chart" viewBox="0 0 420 150" preserveAspectRatio="none"></svg>
            </div>
            <div class="hero-trend-bottom">
              <div class="hero-trend-axis">
                ${trendLabelMonths.map((label) => `<span>${label}</span>`).join("")}
              </div>
              <div class="hero-trend-legend-inline" aria-label="图例">
                <div class="hero-trend-legend-item">
                  <span class="hero-trend-legend-marker hero-trend-legend-marker-bar" aria-hidden="true"></span>
                  <span class="hero-trend-legend-label">支出</span>
                </div>
                <div class="hero-trend-legend-item">
                  <span class="hero-trend-legend-marker hero-trend-legend-marker-line" aria-hidden="true">
                    <span class="hero-trend-legend-marker-line-dot"></span>
                  </span>
                  <span class="hero-trend-legend-label">HC</span>
                </div>
              </div>
            </div>
          </div>
        </article>
      </div>
    `;
  }

  renderHistoricalBreakdownPanel(monthlySeries);
  renderHistoricalTrendChart(monthlySeries);
  return overview;
}

function renderStep1WhatIf() {
  const overview = getHistoricalOverviewMetrics(buildHistoricalMonthlySeries());
  const budgetAlert = getBudgetAlertMetrics(overview);
  const extraCost = budgetAlert.projectedMonthlyCost * state.delayMonths * (1 + state.delayHcGrowthRate);
  const delayedTotalCost = budgetAlert.cumulativeProjectedCost + extraCost;
  const paybackPeriodRdCost = budgetAlert.projectedMonthlyCost * state.paybackMonths;
  const requiredTotalRevenue = state.grossMarginRate > 0 ? paybackPeriodRdCost / state.grossMarginRate : 0;
  const requiredCumulativeTotalRevenue = state.grossMarginRate > 0
    ? (delayedTotalCost + paybackPeriodRdCost) / state.grossMarginRate
    : 0;
  const requiredMonthlyRevenue = state.paybackMonths > 0 ? requiredCumulativeTotalRevenue / state.paybackMonths : 0;

  if (els.delayMonthsRange) els.delayMonthsRange.value = String(state.delayMonths);
  if (els.delayHcGrowthRange) els.delayHcGrowthRange.value = String(Math.round(state.delayHcGrowthRate * 100));
  if (els.paybackMonthsRange) els.paybackMonthsRange.value = String(state.paybackMonths);
  if (els.grossMarginRange) els.grossMarginRange.value = String(Math.round(state.grossMarginRate * 100));
  if (els.delayMonthsValue) els.delayMonthsValue.textContent = `${state.delayMonths} 个月`;
  if (els.delayHcGrowthValue) els.delayHcGrowthValue.textContent = `${state.delayHcGrowthRate > 0 ? "+" : ""}${Math.round(state.delayHcGrowthRate * 100)}%`;
  if (els.paybackMonthsValue) els.paybackMonthsValue.textContent = `${state.paybackMonths} 个月`;
  if (els.grossMarginValue) els.grossMarginValue.textContent = `${Math.round(state.grossMarginRate * 100)}%`;
  if (els.delayExtraCost) els.delayExtraCost.textContent = formatWan(extraCost);
  if (els.delayTotalCost) els.delayTotalCost.textContent = formatWan(delayedTotalCost);
  if (els.requiredTotalRevenue) els.requiredTotalRevenue.textContent = formatWan(requiredTotalRevenue);
  if (els.requiredCumulativeTotalRevenue) els.requiredCumulativeTotalRevenue.textContent = formatWan(requiredCumulativeTotalRevenue);
  if (els.requiredMonthlyRevenue) els.requiredMonthlyRevenue.textContent = formatWanPerMonth(requiredMonthlyRevenue);
}

function renderStep1PivotPanel() {
  if (els.pivotPanel) {
    els.pivotPanel.classList.toggle("hidden", !state.pivotExpanded);
  }
  if (els.pivotToggleButton) {
    els.pivotToggleButton.textContent = state.pivotExpanded ? "收起损益表" : "展开损益表";
  }
}

function renderPostPivot() {
  const pivot = state.postPivot;
  els.pivotDisplayName.textContent = pivot.displayName || "-";
  els.pivotMaxDt.textContent = pivot.maxDt || "-";
  els.pivotHead.innerHTML = `
    <tr>
      <th class="pivot-sticky pivot-code">科目代码</th>
      <th class="pivot-sticky pivot-name">科目名称</th>
      ${pivot.dtColumns.map((dt) => `<th>${dt}</th>`).join("")}
    </tr>
  `;

  if (!pivot.rows.length) {
    els.pivotBody.innerHTML = `<tr><td colspan="${2 + Math.max(pivot.dtColumns.length, 1)}" class="empty-state">暂无分摊后透视数据</td></tr>`;
    return;
  }

  els.pivotBody.innerHTML = pivot.rows.map((row) => `
    <tr>
      <td class="pivot-sticky pivot-code">${row.code}</td>
      <td class="pivot-sticky pivot-name">${formatPivotName(row)}</td>
      ${pivot.dtColumns.map((dt) => `<td class="pivot-value">${formatPivotValue(row.values[dt])}</td>`).join("")}
    </tr>
  `).join("");
}

function renderHcModule() {
  const months = getHcMonths();
  const hcRows = state.hcForecastRows.map((row, accountIndex) => ({
    ...row,
    accountIndex
  }));
  const activeCell = getActiveHcCellMeta();
  const activeRawValue = activeCell.code && activeCell.month ? getHcCellRawValue(activeCell.code, activeCell.month) : "";
  els.hcModule.innerHTML = `
    <div class="hc-excel-head">
      <div class="hc-module-note">步骤 2.1 已按 Excel 模板识别：仅首行标记为“默认值，可点选调整”的字段允许调整，其余字段按模板规则自动识别。</div>
      <div class="hc-module-note">当前预测区间：${formatMonthSlashLabel(state.forecastStart)} 起，共 ${state.forecastMonths} 个月；费用类型为“计算项”的行会按预估方法公式自动汇总。</div>
    </div>
    <div class="hc-formula-bar">
      <div class="hc-name-box">${getHcCellAddress(activeCell.rowIndex, activeCell.colIndex)}</div>
      <label class="hc-formula-field">
        <span>fx</span>
        <input id="hc-formula-input" type="text" value="${activeRawValue}" placeholder="输入数值或公式，例如 =A1.1@2026/5+15" ${getHcRowByCode(activeCell.code)?.computed ? "readonly" : ""} />
      </label>
    </div>
    <div class="hc-shortcuts">
      <button class="hc-action-btn" type="button" data-hc-action="fill-to-end">向右填充到末月</button>
      <span>Enter 下移到下一科目</span>
      <span>Shift+Enter 上移到上一科目</span>
      <span>Tab 横向切换月份</span>
      <span>Ctrl/Cmd+D 向右填充</span>
      <span>支持从 Excel 直接整行粘贴月份数据</span>
      <span>跨月公式写法：=A1.1@2026/5+15</span>
      <span>输入 = 后点击目标单元格，可自动插入引用</span>
      <span>向右拖动填充时，@月份引用会自动顺延</span>
      <span>双击填充柄，可自动补到最后一个预测月</span>
    </div>
    <div class="table-wrap wide hc-excel-wrap">
      <table class="hc-excel-table">
        <thead>
          <tr class="hc-group-row">
            <th class="hc-locked-col hc-step-col" rowspan="2">步骤</th>
            <th class="hc-locked-col hc-code-col" rowspan="2">Code</th>
            <th class="hc-locked-col hc-name-col" rowspan="2">科目名称</th>
            <th class="hc-type-col" rowspan="2">费用类型</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell"></th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            ${months.map(() => `<th class="hc-group-cell is-month">月预估值</th>`).join("")}
          </tr>
          <tr class="hc-subhead-row">
            <th class="hc-method-col">预估方法与参数</th>
            <th class="hc-driver-col">成本动因</th>
            <th class="hc-suggest-logic-col">建议参数（历史值）计算逻辑</th>
            <th class="hc-reference-period-col">建议参数（历史值）-参考期间</th>
            <th class="hc-suggest-value-col">建议参数（历史值）</th>
            <th class="hc-forecast-param-col">预估参数</th>
            <th class="hc-forecast-logic-col">预估值计算逻辑</th>
            ${months.map((month) => `<th class="hc-month-header">${formatMonthSlashLabel(month)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${hcRows.map((row) => `
            <tr class="${row.computed ? "hc-summary-row hc-row-auto" : "hc-row-manual"}">
              <td class="hc-locked-col hc-step-col" data-hc-row-header="${row.accountIndex}">${hcStepLabel}</td>
              <td class="hc-locked-col hc-code-col" data-hc-row-header="${row.accountIndex}">${row.code}</td>
              <td class="hc-locked-col hc-name-col" data-hc-row-header="${row.accountIndex}">${row.account}</td>
              <td class="hc-meta-text">${row.feeType || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.estimateMethodRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${hcMetaFieldListIds.estimateMethodRaw}" data-hc-meta-field="estimateMethodRaw" data-hc-meta-code="${row.code}" type="text" value="${escapeHtml(row.estimateMethodRaw || "")}" />`}
              </td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.costDriver || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${hcMetaFieldListIds.costDriver}" data-hc-meta-field="costDriver" data-hc-meta-code="${row.code}" type="text" value="${escapeHtml(row.costDriver || "")}" />`}
              </td>
              <td class="hc-meta-text">${row.suggestedLogic || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.referencePeriodRaw || "-"}</div>`
                  : `<select class="cell-select hc-meta-select" data-hc-meta-field="referencePeriodRaw" data-hc-meta-code="${row.code}">
                      ${hcMetaOptionSets.referencePeriodRaw.map((option) => `<option value="${option}" ${row.referencePeriodRaw === option ? "selected" : ""}>${option}</option>`).join("")}
                    </select>`}
              </td>
              <td class="hc-meta-text">${row.suggestedValueText || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.forecastParamRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" data-hc-meta-field="forecastParamRaw" data-hc-meta-code="${row.code}" type="text" value="${escapeHtml(row.forecastParamRaw || "")}" placeholder="${row.estimateMethodRaw === "人均成本" ? "例如 15" : row.estimateMethodRaw === "月费用额" ? "例如 20" : ""}" />`}
              </td>
              <td class="hc-meta-text">${row.forecastLogicRaw || "-"}</td>
              ${months.map((month, monthIndex) => {
                const rawValue = getHcCellRawValue(row.code, month);
                const numericValue = getHcCellNumericValue(row.code, month);
                const logicalRowIndex = row.accountIndex;
                const logicalColIndex = monthIndex;
                const isEditing = isSameHcCell(state.hcEditCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex });
                const displayValue = isEditing ? rawValue : (rawValue === "" ? "" : compactDecimal(numericValue, 2));
                const editingClass = isEditing ? "is-editing" : "";
                const selectedClass = isHcCellSelected(logicalRowIndex, logicalColIndex) ? "is-selected" : "";
                const activeClass = isSameHcCell(activeCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex }) ? "is-active" : "";
                const previewClass = state.hcFillDrag && logicalColIndex > normalizeHcSelection().endCol && logicalColIndex <= state.hcFillDrag.targetCol && logicalRowIndex >= normalizeHcSelection().startRow && logicalRowIndex <= normalizeHcSelection().endRow ? "is-fill-preview" : "";
                const tailCell = !row.computed
                  ? `<button class="hc-fill-handle" type="button" data-hc-fill-handle="true" tabindex="-1" aria-label="向右拖动填充，双击直达末月" title="向右拖动填充，双击直达末月"></button>`
                  : "";
                return `
                <td class="hc-cell ${editingClass} ${selectedClass} ${activeClass} ${previewClass}" data-hc-row="${logicalRowIndex}" data-hc-col="${logicalColIndex}">
                  <input
                    class="cell-input hc-cell-input"
                    data-hc-code="${row.code}"
                    data-hc-month="${month}"
                    data-hc-row="${logicalRowIndex}"
                    data-hc-col="${logicalColIndex}"
                    type="text"
                    inputmode="decimal"
                    value="${displayValue}"
                    ${row.computed ? "readonly" : ""}
                  />
                  ${tailCell}
                </td>
              `;
              }).join("")}
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
    ${renderHcMetaDatalist(hcMetaFieldListIds.estimateMethodRaw, hcMetaOptionSets.estimateMethodRaw)}
    ${renderHcMetaDatalist(hcMetaFieldListIds.costDriver, hcMetaOptionSets.costDriver)}
  `;
}

function renderRevenueModule() {
  const months = getRevenueMonths();
  const summaryRows = getRevenueForecastSummary(months);
  const naturalYearRows = buildRevenueNaturalYearSummaries(summaryRows);
  const glYearRows = buildRevenueGlYearSummaries(summaryRows, state.goLiveDate);
  const revenueRows = state.revenueForecastRows.map((row, accountIndex) => ({
    ...row,
    accountIndex
  }));
  const importedFlowNote = state.flowImportName
    ? `A2 Gross Revenue 已从步骤 1 带入：${state.flowImportName}（${state.flowImportRowCount} 个月）；A3 / A4 会按费用率和 A2 自动联动，A5 按公式自动汇总。`
    : "当前使用演示流水作为 A2 默认值；导入步骤 1 的流水 Excel 后，会自动回填到这里，并保留后续手动调整能力。";
  const activeCell = getActiveRevenueCellMeta();
  const activeRawValue = activeCell.code && activeCell.month ? getRevenueCellEditValue(activeCell.code, activeCell.month) : "";

  els.revenueModule.innerHTML = `
    <div class="hc-excel-head">
      <div class="hc-module-note">步骤 2.2 已按 Excel 模板识别：A2/A3/A4/A5 分别对应 Gross Revenue、VAT、Platform Cost、Net Revenue。</div>
      <div class="hc-module-note">当前预测区间：${formatMonthSlashLabel(state.forecastStart)} 起，共 ${state.forecastMonths} 个月；金额展示统一为整数 USD，费用类型为“计算项”的 A5 会按预估方法公式自动汇总。</div>
    </div>
    <div class="revenue-settings-bar">
      <label class="draft-slot-field revenue-setting-field">
        <span>GL 时间</span>
        <input type="date" data-revenue-setting="goLiveDate" value="${escapeHtml(state.goLiveDate || "")}" />
      </label>
      <div class="hc-module-note">结果页会基于这里的 GL 时间生成 GL1Y - GL5Y 汇总。</div>
    </div>
    <div class="hc-formula-bar revenue-formula-bar">
      <div class="hc-name-box">${getRevenueCellAddress(activeCell.rowIndex, activeCell.colIndex)}</div>
      <label class="hc-formula-field">
        <span>fx</span>
        <input id="revenue-formula-input" type="text" value="${activeRawValue}" placeholder="输入数值或公式，例如 =A2@2026/5*6%" ${isAutoCalculatedForecastRow(getRevenueRowByCode(activeCell.code)) ? "readonly" : ""} />
      </label>
    </div>
    <div class="hc-shortcuts">
      <button class="hc-action-btn revenue-action-btn" type="button" data-revenue-action="fill-to-end">向右填充到末月</button>
      <span>Enter 下移到下一科目</span>
      <span>Shift+Enter 上移到上一科目</span>
      <span>Tab 横向切换月份</span>
      <span>Ctrl/Cmd+D 向右填充</span>
      <span>支持从 Excel 直接整行粘贴月份数据</span>
      <span>跨月公式写法：=A2@2026/5+10000</span>
      <span>输入 = 后点击目标单元格，可自动插入引用</span>
      <span>双击填充柄，可自动补到最后一个预测月</span>
    </div>
    <div class="revenue-import-bar">
      <div class="hc-module-note revenue-import-note">${importedFlowNote}</div>
      ${state.flowImportName ? `<button class="hc-action-btn revenue-action-btn" type="button" data-revenue-action="restore-imported-flow">按导入流水回填</button>` : ""}
    </div>
    <div class="table-wrap wide hc-excel-wrap">
      <table class="hc-excel-table revenue-input-table">
        <thead>
          <tr class="hc-group-row">
            <th class="hc-locked-col hc-step-col" rowspan="2">步骤</th>
            <th class="hc-locked-col hc-code-col" rowspan="2">Code</th>
            <th class="hc-locked-col hc-name-col" rowspan="2">科目名称</th>
            <th class="hc-type-col" rowspan="2">费用类型</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell"></th>
            <th class="hc-group-cell is-editable">默认值，可手工调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            ${months.map(() => `<th class="hc-group-cell is-month">月预估值</th>`).join("")}
          </tr>
          <tr class="hc-subhead-row">
            <th class="hc-method-col">预估方法与参数</th>
            <th class="hc-driver-col">成本动因</th>
            <th class="hc-suggest-logic-col">建议参数（历史值）计算逻辑</th>
            <th class="hc-reference-period-col">建议参数（历史值）-参考期间</th>
            <th class="hc-suggest-value-col">建议参数（历史值）</th>
            <th class="hc-forecast-param-col">预估参数</th>
            <th class="hc-forecast-logic-col">预估值计算逻辑</th>
            ${months.map((month) => `<th class="hc-month-header">${formatMonthSlashLabel(month)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${revenueRows.map((row) => {
            const rowToneClass = isAutoCalculatedForecastRow(row) ? "hc-row-auto" : "hc-row-manual";
            return `
            <tr class="${row.computed ? "hc-summary-row " : ""}${rowToneClass}">
              <td class="hc-locked-col hc-step-col" data-revenue-row-header="${row.accountIndex}">2.2</td>
              <td class="hc-locked-col hc-code-col" data-revenue-row-header="${row.accountIndex}">${row.code}</td>
              <td class="hc-locked-col hc-name-col" data-revenue-row-header="${row.accountIndex}">${row.account}</td>
              <td class="hc-meta-text">${row.feeType || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.estimateMethodRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${revenueMetaFieldListIds.estimateMethodRaw}" data-revenue-meta-field="estimateMethodRaw" data-revenue-meta-code="${row.code}" type="text" value="${escapeHtml(row.estimateMethodRaw || "")}" />`}
              </td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.costDriver || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${revenueMetaFieldListIds.costDriver}" data-revenue-meta-field="costDriver" data-revenue-meta-code="${row.code}" type="text" value="${escapeHtml(row.costDriver || "")}" />`}
              </td>
              <td class="hc-meta-text">${row.suggestedLogic || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.referencePeriodRaw || "-"}</div>`
                  : `<select class="cell-select hc-meta-select" data-revenue-meta-field="referencePeriodRaw" data-revenue-meta-code="${row.code}">
                      ${revenueMetaOptionSets.referencePeriodRaw.map((option) => `<option value="${option}" ${row.referencePeriodRaw === option ? "selected" : ""}>${option}</option>`).join("")}
                    </select>`}
              </td>
              <td class="hc-meta-text">${row.suggestedValueText || "-"}</td>
              <td>
                ${shouldLockForecastParam(row)
                  ? `<div class="hc-meta-text">${row.forecastParamRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" data-revenue-meta-field="forecastParamRaw" data-revenue-meta-code="${row.code}" type="text" value="${escapeHtml(row.forecastParamRaw || "")}" placeholder="${row.estimateMethodRaw === "费用率" ? "例如 6%" : ""}" />`}
              </td>
              <td class="hc-meta-text">${row.forecastLogicRaw || "-"}</td>
              ${months.map((month, monthIndex) => {
                const rawValue = getRevenueCellRawValue(row.code, month);
                const numericValue = getRevenueCellNumericValue(row.code, month);
                const logicalRowIndex = row.accountIndex;
                const logicalColIndex = monthIndex;
                const autoCalculated = isAutoCalculatedForecastRow(row);
                const isEditing = isSameHcCell(state.revenueEditCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex });
                const displayValue = isEditing ? getRevenueCellEditValue(row.code, month) : ((rawValue !== "" || numericValue) ? compactDecimal(numericValue, 0) : "");
                const editingClass = isEditing ? "is-editing" : "";
                const selectedClass = isRevenueCellSelected(logicalRowIndex, logicalColIndex) ? "is-selected" : "";
                const activeClass = isSameHcCell(activeCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex }) ? "is-active" : "";
                const previewClass = state.revenueFillDrag
                  && logicalColIndex > normalizeRevenueSelection().endCol
                  && logicalColIndex <= state.revenueFillDrag.targetCol
                  && logicalRowIndex >= normalizeRevenueSelection().startRow
                  && logicalRowIndex <= normalizeRevenueSelection().endRow
                  ? "is-fill-preview"
                  : "";
                const tailCell = !autoCalculated
                  ? `<button class="hc-fill-handle" type="button" data-revenue-fill-handle="true" tabindex="-1" aria-label="向右拖动填充，双击直达末月" title="向右拖动填充，双击直达末月"></button>`
                  : "";
                return `
                  <td class="hc-cell revenue-grid-cell ${editingClass} ${selectedClass} ${activeClass} ${previewClass}" data-revenue-row="${logicalRowIndex}" data-revenue-col="${logicalColIndex}">
                    <input
                      class="cell-input revenue-cell-input"
                      data-revenue-code="${row.code}"
                      data-revenue-month="${month}"
                      data-revenue-row="${logicalRowIndex}"
                      data-revenue-col="${logicalColIndex}"
                      type="text"
                      inputmode="decimal"
                      value="${displayValue}"
                      ${autoCalculated ? "readonly" : ""}
                    />
                    ${tailCell}
                  </td>
                `;
              }).join("")}
            </tr>
          `;
          }).join("")}
        </tbody>
      </table>
    </div>
    <div class="revenue-annual-summary">
      <div class="section-title">
        <h3>年度汇总</h3>
        <p>此处展示年度口径汇总：包括自然年汇总，以及按 GL 时间切分的 GL1 - GL5；GL 年和自然年一样，未满 12 个月时也会按当前已覆盖月份汇总展示。</p>
      </div>
      <div class="revenue-summary-grid">
        <div class="table-wrap revenue-summary-wrap">
          <table class="rules-table revenue-summary-table revenue-annual-table">
            <thead>
              <tr>
                <th>自然年</th>
                <th>覆盖月份</th>
                <th>A2 项目流水</th>
                <th>A3 增值税</th>
                <th>A4 平台费</th>
                <th>A5 净收入</th>
              </tr>
            </thead>
            <tbody>
              ${renderRevenueAnnualSummaryTableRows(naturalYearRows, "当前预测区间内暂无可汇总的自然年数据。")}
            </tbody>
          </table>
        </div>
        <div class="table-wrap revenue-summary-wrap">
          <table class="rules-table revenue-summary-table revenue-annual-table">
            <thead>
              <tr>
                <th>GL 年</th>
                <th>覆盖月份</th>
                <th>A2 项目流水</th>
                <th>A3 增值税</th>
                <th>A4 平台费</th>
                <th>A5 净收入</th>
              </tr>
            </thead>
            <tbody>
              ${renderRevenueAnnualSummaryTableRows(glYearRows, state.goLiveDate ? "当前预测区间内暂无可汇总的 GL1 - GL5 数据。" : "请先填写 GL 时间后查看 GL1 - GL5 汇总。")}
            </tbody>
          </table>
        </div>
      </div>
    </div>
    ${renderHcMetaDatalist(revenueMetaFieldListIds.estimateMethodRaw, revenueMetaOptionSets.estimateMethodRaw)}
    ${renderHcMetaDatalist(revenueMetaFieldListIds.costDriver, revenueMetaOptionSets.costDriver)}
  `;
}

function renderCostModule() {
  const months = getCostMonths();
  const costRows = state.costForecastRows.map((row, accountIndex) => ({ ...row, accountIndex }));
  const activeCell = getActiveCostCellMeta();
  const activeRawValue = activeCell.code && activeCell.month ? getCostCellEditValue(activeCell.code, activeCell.month) : "";
  els.costModule.innerHTML = `
    <div class="hc-excel-head">
      <div class="hc-module-note">步骤 2.3 已按 Excel 模板识别：A97 / A91 为收入项，A6 / A7 / A8 / A54 为费用项，A9 为计算项。</div>
      <div class="hc-module-note">当前预测区间：${formatMonthSlashLabel(state.forecastStart)} 起，共 ${state.forecastMonths} 个月；A9 会按 =A5+A97+A91-A6-A7-A8-A54 自动汇总。</div>
    </div>
    <div class="hc-formula-bar revenue-formula-bar">
      <div class="hc-name-box">${getCostCellAddress(activeCell.rowIndex, activeCell.colIndex)}</div>
      <label class="hc-formula-field">
        <span>fx</span>
        <input id="cost-formula-input" type="text" value="${activeRawValue}" placeholder="输入数值或公式，例如 =A97@2026/5+1000" ${isAutoCalculatedForecastRow(getCostRowByCode(activeCell.code)) ? "readonly" : ""} />
      </label>
    </div>
    <div class="hc-shortcuts">
      <button class="hc-action-btn revenue-action-btn" type="button" data-cost-action="fill-to-end">向右填充到末月</button>
      <span>Enter 下移到下一科目</span>
      <span>Shift+Enter 上移到上一科目</span>
      <span>Tab 横向切换月份</span>
      <span>Ctrl/Cmd+D 向右填充</span>
      <span>支持从 Excel 直接整行粘贴月份数据</span>
      <span>跨月公式写法：=A97@2026/5+1000</span>
      <span>输入 = 后点击目标单元格，可自动插入引用</span>
      <span>双击填充柄，可自动补到最后一个预测月</span>
    </div>
    <div class="table-wrap wide hc-excel-wrap">
      <table class="hc-excel-table revenue-input-table">
        <thead>
          <tr class="hc-group-row">
            <th class="hc-locked-col hc-step-col" rowspan="2">步骤</th>
            <th class="hc-locked-col hc-code-col" rowspan="2">Code</th>
            <th class="hc-locked-col hc-name-col" rowspan="2">科目名称</th>
            <th class="hc-type-col" rowspan="2">费用类型</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell"></th>
            <th class="hc-group-cell is-editable">默认值，可手工调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            ${months.map(() => `<th class="hc-group-cell is-month">月预估值</th>`).join("")}
          </tr>
          <tr class="hc-subhead-row">
            <th class="hc-method-col">预估方法与参数</th>
            <th class="hc-driver-col">成本动因</th>
            <th class="hc-suggest-logic-col">建议参数（历史值）计算逻辑</th>
            <th class="hc-reference-period-col">建议参数（历史值）-参考期间</th>
            <th class="hc-suggest-value-col">建议参数（历史值）</th>
            <th class="hc-forecast-param-col">预估参数</th>
            <th class="hc-forecast-logic-col">预估值计算逻辑</th>
            ${months.map((month) => `<th class="hc-month-header">${formatMonthSlashLabel(month)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${costRows.map((row) => {
            const rowToneClass = isAutoCalculatedForecastRow(row) ? "hc-row-auto" : "hc-row-manual";
            return `
            <tr class="${row.computed ? "hc-summary-row " : ""}${rowToneClass}">
              <td class="hc-locked-col hc-step-col" data-cost-row-header="${row.accountIndex}">${costStepLabel}</td>
              <td class="hc-locked-col hc-code-col" data-cost-row-header="${row.accountIndex}">${row.code}</td>
              <td class="hc-locked-col hc-name-col" data-cost-row-header="${row.accountIndex}">${row.account}</td>
              <td class="hc-meta-text">${row.feeType || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.estimateMethodRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${costMetaFieldListIds.estimateMethodRaw}" data-cost-meta-field="estimateMethodRaw" data-cost-meta-code="${row.code}" type="text" value="${escapeHtml(row.estimateMethodRaw || "")}" />`}
              </td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.costDriver || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${costMetaFieldListIds.costDriver}" data-cost-meta-field="costDriver" data-cost-meta-code="${row.code}" type="text" value="${escapeHtml(row.costDriver || "")}" />`}
              </td>
              <td class="hc-meta-text">${row.suggestedLogic || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.referencePeriodRaw || "-"}</div>`
                  : `<select class="cell-select hc-meta-select" data-cost-meta-field="referencePeriodRaw" data-cost-meta-code="${row.code}">
                    ${costMetaOptionSets.referencePeriodRaw.map((option) => `<option value="${option}" ${row.referencePeriodRaw === option ? "selected" : ""}>${option}</option>`).join("")}
                  </select>`}
              </td>
              <td class="hc-meta-text">${row.suggestedValueText || "-"}</td>
              <td>
                ${shouldLockForecastParam(row)
                  ? `<div class="hc-meta-text">${row.forecastParamRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" data-cost-meta-field="forecastParamRaw" data-cost-meta-code="${row.code}" type="text" value="${escapeHtml(row.forecastParamRaw || "")}" placeholder="${row.estimateMethodRaw === "费用率" ? "例如 6%" : ""}" />`}
              </td>
              <td class="hc-meta-text">${row.forecastLogicRaw || "-"}</td>
              ${months.map((month, monthIndex) => {
                const rawValue = getCostCellRawValue(row.code, month);
                const numericValue = getCostCellNumericValue(row.code, month);
                const logicalRowIndex = row.accountIndex;
                const logicalColIndex = monthIndex;
                const autoCalculated = isAutoCalculatedForecastRow(row);
                const isEditing = isSameHcCell(state.costEditCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex });
                const displayValue = isEditing ? getCostCellEditValue(row.code, month) : ((rawValue !== "" || numericValue) ? compactDecimal(numericValue, 0) : "");
                const editingClass = isEditing ? "is-editing" : "";
                const selectedClass = isCostCellSelected(logicalRowIndex, logicalColIndex) ? "is-selected" : "";
                const activeClass = isSameHcCell(activeCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex }) ? "is-active" : "";
                const previewClass = state.costFillDrag
                  && logicalColIndex > normalizeCostSelection().endCol
                  && logicalColIndex <= state.costFillDrag.targetCol
                  && logicalRowIndex >= normalizeCostSelection().startRow
                  && logicalRowIndex <= normalizeCostSelection().endRow
                  ? "is-fill-preview" : "";
                const tailCell = !autoCalculated
                  ? `<button class="hc-fill-handle" type="button" data-cost-fill-handle="true" tabindex="-1" aria-label="向右拖动填充，双击直达末月" title="向右拖动填充，双击直达末月"></button>`
                  : "";
                return `
                  <td class="hc-cell revenue-grid-cell ${editingClass} ${selectedClass} ${activeClass} ${previewClass}" data-cost-row="${logicalRowIndex}" data-cost-col="${logicalColIndex}">
                    <input
                      class="cell-input revenue-cell-input"
                      data-cost-code="${row.code}"
                      data-cost-month="${month}"
                      data-cost-row="${logicalRowIndex}"
                      data-cost-col="${logicalColIndex}"
                      type="text"
                      inputmode="decimal"
                      value="${displayValue}"
                      ${autoCalculated ? "readonly" : ""}
                    />
                    ${tailCell}
                  </td>
                `;
              }).join("")}
            </tr>
          `;
          }).join("")}
        </tbody>
      </table>
    </div>
    ${renderHcMetaDatalist(costMetaFieldListIds.estimateMethodRaw, costMetaOptionSets.estimateMethodRaw)}
    ${renderHcMetaDatalist(costMetaFieldListIds.costDriver, costMetaOptionSets.costDriver)}
  `;
}

function renderMarketingModule() {
  const months = getMarketingMonths();
  const marketingRows = state.marketingForecastRows.map((row, accountIndex) => ({ ...row, accountIndex }));
  const activeCell = getActiveMarketingCellMeta();
  const activeRawValue = activeCell.code && activeCell.month ? getMarketingCellEditValue(activeCell.code, activeCell.month) : "";
  els.marketingModule.innerHTML = `
    <div class="hc-excel-head">
      <div class="hc-module-note">步骤 2.4 已按 Excel 模板识别：A11 / A11.1 / A11.2 为市场费率链路，A12 / A12.1 为手工输入项，A10 / A99 为计算项。</div>
      <div class="hc-module-note">当前预测区间：${formatMonthSlashLabel(state.forecastStart)} 起，共 ${state.forecastMonths} 个月；A10 会按 =A11+A11.1+A11.2+A12+A12.1 自动汇总，A99 会按 =A9-A10 自动计算。</div>
    </div>
    <div class="hc-formula-bar revenue-formula-bar">
      <div class="hc-name-box">${getMarketingCellAddress(activeCell.rowIndex, activeCell.colIndex)}</div>
      <label class="hc-formula-field">
        <span>fx</span>
        <input id="marketing-formula-input" type="text" value="${activeRawValue}" placeholder="输入数值或公式，例如 =A11@2026/5*10%" ${isAutoCalculatedForecastRow(getMarketingRowByCode(activeCell.code)) ? "readonly" : ""} />
      </label>
    </div>
    <div class="hc-shortcuts">
      <button class="hc-action-btn revenue-action-btn" type="button" data-marketing-action="fill-to-end">向右填充到末月</button>
      <span>Enter 下移到下一科目</span>
      <span>Shift+Enter 上移到上一科目</span>
      <span>Tab 横向切换月份</span>
      <span>Ctrl/Cmd+D 向右填充</span>
      <span>支持从 Excel 直接整行粘贴月份数据</span>
      <span>跨月公式写法：=A11@2026/5*10%</span>
      <span>输入 = 后点击目标单元格，可自动插入引用</span>
      <span>双击填充柄，可自动补到最后一个预测月</span>
    </div>
    <div class="table-wrap wide hc-excel-wrap">
      <table class="hc-excel-table revenue-input-table">
        <thead>
          <tr class="hc-group-row">
            <th class="hc-locked-col hc-step-col" rowspan="2">步骤</th>
            <th class="hc-locked-col hc-code-col" rowspan="2">Code</th>
            <th class="hc-locked-col hc-name-col" rowspan="2">科目名称</th>
            <th class="hc-type-col" rowspan="2">费用类型</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            <th class="hc-group-cell is-editable">默认值，可点选调整</th>
            <th class="hc-group-cell"></th>
            <th class="hc-group-cell is-editable">默认值，可手工调整</th>
            <th class="hc-group-cell is-derived">默认值，不可点选，根据选定的预估方法与参数映射</th>
            ${months.map(() => `<th class="hc-group-cell is-month">月预估值</th>`).join("")}
          </tr>
          <tr class="hc-subhead-row">
            <th class="hc-method-col">预估方法与参数</th>
            <th class="hc-driver-col">成本动因</th>
            <th class="hc-suggest-logic-col">建议参数（历史值）计算逻辑</th>
            <th class="hc-reference-period-col">建议参数（历史值）-参考期间</th>
            <th class="hc-suggest-value-col">建议参数（历史值）</th>
            <th class="hc-forecast-param-col">预估参数</th>
            <th class="hc-forecast-logic-col">预估值计算逻辑</th>
            ${months.map((month) => `<th class="hc-month-header">${formatMonthSlashLabel(month)}</th>`).join("")}
          </tr>
        </thead>
        <tbody>
          ${marketingRows.map((row) => {
            const rowToneClass = isAutoCalculatedForecastRow(row) ? "hc-row-auto" : "hc-row-manual";
            return `
            <tr class="${row.computed ? "hc-summary-row " : ""}${rowToneClass}">
              <td class="hc-locked-col hc-step-col" data-marketing-row-header="${row.accountIndex}">${marketingStepLabel}</td>
              <td class="hc-locked-col hc-code-col" data-marketing-row-header="${row.accountIndex}">${row.code}</td>
              <td class="hc-locked-col hc-name-col" data-marketing-row-header="${row.accountIndex}">${row.account}</td>
              <td class="hc-meta-text">${row.feeType || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.estimateMethodRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${marketingMetaFieldListIds.estimateMethodRaw}" data-marketing-meta-field="estimateMethodRaw" data-marketing-meta-code="${row.code}" type="text" value="${escapeHtml(row.estimateMethodRaw || "")}" />`}
              </td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.costDriver || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" list="${marketingMetaFieldListIds.costDriver}" data-marketing-meta-field="costDriver" data-marketing-meta-code="${row.code}" type="text" value="${escapeHtml(row.costDriver || "")}" />`}
              </td>
              <td class="hc-meta-text">${row.suggestedLogic || "-"}</td>
              <td>
                ${row.computed
                  ? `<div class="hc-meta-text">${row.referencePeriodRaw || "-"}</div>`
                  : `<select class="cell-select hc-meta-select" data-marketing-meta-field="referencePeriodRaw" data-marketing-meta-code="${row.code}">
                    ${marketingMetaOptionSets.referencePeriodRaw.map((option) => `<option value="${option}" ${row.referencePeriodRaw === option ? "selected" : ""}>${option}</option>`).join("")}
                  </select>`}
              </td>
              <td class="hc-meta-text">${row.suggestedValueText || "-"}</td>
              <td>
                ${shouldLockForecastParam(row)
                  ? `<div class="hc-meta-text">${row.forecastParamRaw || "-"}</div>`
                  : `<input class="cell-input hc-meta-input" data-marketing-meta-field="forecastParamRaw" data-marketing-meta-code="${row.code}" type="text" value="${escapeHtml(row.forecastParamRaw || "")}" placeholder="${row.estimateMethodRaw === "费用率" ? "例如 10%" : ""}" />`}
              </td>
              <td class="hc-meta-text">${row.forecastLogicRaw || "-"}</td>
              ${months.map((month, monthIndex) => {
                const rawValue = getMarketingCellRawValue(row.code, month);
                const numericValue = getMarketingCellNumericValue(row.code, month);
                const logicalRowIndex = row.accountIndex;
                const logicalColIndex = monthIndex;
                const autoCalculated = isAutoCalculatedForecastRow(row);
                const isEditing = isSameHcCell(state.marketingEditCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex });
                const displayValue = isEditing ? getMarketingCellEditValue(row.code, month) : ((rawValue !== "" || numericValue) ? compactDecimal(numericValue, 0) : "");
                const editingClass = isEditing ? "is-editing" : "";
                const selectedClass = isMarketingCellSelected(logicalRowIndex, logicalColIndex) ? "is-selected" : "";
                const activeClass = isSameHcCell(activeCell, { rowIndex: logicalRowIndex, colIndex: logicalColIndex }) ? "is-active" : "";
                const previewClass = state.marketingFillDrag
                  && logicalColIndex > normalizeMarketingSelection().endCol
                  && logicalColIndex <= state.marketingFillDrag.targetCol
                  && logicalRowIndex >= normalizeMarketingSelection().startRow
                  && logicalRowIndex <= normalizeMarketingSelection().endRow
                  ? "is-fill-preview" : "";
                const tailCell = !autoCalculated
                  ? `<button class="hc-fill-handle" type="button" data-marketing-fill-handle="true" tabindex="-1" aria-label="向右拖动填充，双击直达末月" title="向右拖动填充，双击直达末月"></button>`
                  : "";
                return `
                  <td class="hc-cell revenue-grid-cell ${editingClass} ${selectedClass} ${activeClass} ${previewClass}" data-marketing-row="${logicalRowIndex}" data-marketing-col="${logicalColIndex}">
                    <input
                      class="cell-input revenue-cell-input"
                      data-marketing-code="${row.code}"
                      data-marketing-month="${month}"
                      data-marketing-row="${logicalRowIndex}"
                      data-marketing-col="${logicalColIndex}"
                      type="text"
                      inputmode="decimal"
                      value="${displayValue}"
                      ${autoCalculated ? "readonly" : ""}
                    />
                    ${tailCell}
                  </td>
                `;
              }).join("")}
            </tr>
          `;
          }).join("")}
        </tbody>
      </table>
    </div>
    ${renderHcMetaDatalist(marketingMetaFieldListIds.estimateMethodRaw, marketingMetaOptionSets.estimateMethodRaw)}
    ${renderHcMetaDatalist(marketingMetaFieldListIds.costDriver, marketingMetaOptionSets.costDriver)}
  `;
}

function renderRuleTableRows(entries, monthOptions) {
  return entries.map(({ rule, rowIndex }) => `
    <tr class="${isLinkedRuleCode(rule.code) ? "rules-summary-row" : ""}" data-key="${rule.key}" data-rule-row="${rowIndex}">
      <td>${rule.code || "-"}</td>
      <td>${rule.account}</td>
      <td><span class="pill">${rule.source}</span></td>
      <td><select class="cell-select" data-field="type" ${isRuleEditable(rule, "type") ? "" : "disabled"}><option value="fixed" ${rule.type === "fixed" ? "selected" : ""}>固定费用</option><option value="variable" ${rule.type === "variable" ? "selected" : ""}>变动费用</option></select></td>
      <td>${rule.method === "rate" ? ratio(rule.suggestedValue) : compactMoney(rule.suggestedValue)}</td>
      <td class="rules-grid-cell ${isRuleCellSelected(rowIndex, 0) ? "is-selected" : ""} ${isSameHcCell(state.rulesEditCell, { rowIndex, colIndex: 0 }) ? "is-active" : ""} ${state.rulesFillDrag && rowIndex > normalizeRulesSelection().endRow && rowIndex <= state.rulesFillDrag.targetRow && 0 >= normalizeRulesSelection().startCol && 0 <= normalizeRulesSelection().endCol ? "is-fill-preview" : ""}" data-rule-row="${rowIndex}" data-rule-col="0">
        <input class="cell-input rules-grid-input" data-field="expectedValue" data-rule-row="${rowIndex}" data-rule-col="0" type="text" inputmode="decimal" ${isRuleEditable(rule, "expectedValue") ? "" : "readonly"} value="${isSameHcCell(state.rulesEditCell, { rowIndex, colIndex: 0 }) ? getRuleCellRawValue(rule, "expectedValue") : compactDecimal(rule.expectedValue, 4)}" />
        ${isRuleEditable(rule, "expectedValue") && isRulesSelectionTailCell(rowIndex, 0) ? `<button class="hc-fill-handle" type="button" data-rules-fill-handle="true" tabindex="-1" aria-label="向下填充"></button>` : ""}
      </td>
      <td><select class="cell-select" data-field="driver" ${isRuleEditable(rule, "driver") ? "" : "disabled"}><option value="HC" ${rule.driver === "HC" ? "selected" : ""}>HC</option><option value="流水" ${rule.driver === "流水" ? "selected" : ""}>流水</option></select></td>
      <td><select class="cell-select" data-field="method" ${isRuleEditable(rule, "method") ? "" : "disabled"}><option value="fixed" ${rule.method === "fixed" ? "selected" : ""}>固定额</option><option value="unit" ${rule.method === "unit" ? "selected" : ""}>单位成本</option><option value="rate" ${rule.method === "rate" ? "selected" : ""}>费用率</option></select></td>
      <td><select class="cell-select" data-field="startMonth" ${isRuleEditable(rule, "startMonth") ? "" : "disabled"}>${monthOptions.replace(`value=\"${rule.startMonth}\"`, `value=\"${rule.startMonth}\" selected`)}</select></td>
      <td><select class="cell-select" data-field="endMonth" ${isRuleEditable(rule, "endMonth") ? "" : "disabled"}><option value="tail" ${rule.endMonth === "tail" ? "selected" : ""}>持续至结束</option>${monthOptions.replace(`value=\"${rule.endMonth}\"`, `value=\"${rule.endMonth}\" selected`)}</select></td>
      <td class="rules-grid-cell ${isRuleCellSelected(rowIndex, 1) ? "is-selected" : ""} ${isSameHcCell(state.rulesEditCell, { rowIndex, colIndex: 1 }) ? "is-active" : ""} ${state.rulesFillDrag && rowIndex > normalizeRulesSelection().endRow && rowIndex <= state.rulesFillDrag.targetRow && 1 >= normalizeRulesSelection().startCol && 1 <= normalizeRulesSelection().endCol ? "is-fill-preview" : ""}" data-rule-row="${rowIndex}" data-rule-col="1">
        <input class="cell-input rules-grid-input" data-field="growthRate" data-rule-row="${rowIndex}" data-rule-col="1" type="text" inputmode="decimal" ${isRuleEditable(rule, "growthRate") ? "" : "readonly"} value="${isSameHcCell(state.rulesEditCell, { rowIndex, colIndex: 1 }) ? getRuleCellRawValue(rule, "growthRate") : compactDecimal(rule.growthRate, 4)}" />
        ${isRuleEditable(rule, "growthRate") && isRulesSelectionTailCell(rowIndex, 1) ? `<button class="hc-fill-handle" type="button" data-rules-fill-handle="true" tabindex="-1" aria-label="向下填充"></button>` : ""}
      </td>
      <td class="table-note">${isLinkedRuleCode(rule.code) ? `联动汇总：${getLinkedRuleLabel(rule.code)}` : (rule.type === "fixed" ? "按月固定发生，可设置增长。" : `按${rule.driver}驱动，预估值 = 成本动因 × ${rule.method === "rate" ? "费率" : "单位成本"}`)}</td>
    </tr>
  `).join("");
}

function renderRuleModuleCards(entries, monthOptions) {
  return ruleModuleDefinitions.map((module) => {
    const moduleEntries = entries.filter((entry) => entry.moduleKey === module.key);
    const tableBody = moduleEntries.length ? renderRuleTableRows(moduleEntries, monthOptions) : "";
    const tableHtml = moduleEntries.length
      ? `
        <div class="table-wrap wide">
          <table class="rules-table">
            <thead>
              <tr>
                <th>科目编号</th>
                <th>科目名称</th>
                <th>费用来源</th>
                <th>预估规则</th>
                <th>建议值</th>
                <th>预期值</th>
                <th>成本动因</th>
                <th>计算方法</th>
                <th>起始月</th>
                <th>结束月</th>
                <th>增长设置</th>
                <th>规则说明</th>
              </tr>
            </thead>
            <tbody>${tableBody}</tbody>
          </table>
        </div>
      `
      : `<div class="empty-state">当前筛选条件下，本模块没有匹配科目。</div>`;

    return `
      <section class="rules-module-card">
        <div class="rules-module-head">
          <div class="section-title">
            <h3>${module.title}</h3>
            <p>${module.description}</p>
          </div>
          <div class="module-pill">${moduleEntries.length} 个科目</div>
        </div>
        <div class="rules-module-meta">
          <div class="rules-module-formula">${module.formula}</div>
        </div>
        ${tableHtml}
      </section>
    `;
  }).join("");
}

function renderRules() {
  const months = monthRange(state.forecastStart, state.forecastMonths);
  const monthOptions = months.map((month) => `<option value="${month}">${month}</option>`).join("");
  const scope = getActiveScope();
  const rules = getRulesByScope(scope).filter((rule) => !isManualModuleAccountCode(rule.code));
  const filteredEntries = getRenderableRuleEntries();
  const filtered = filteredEntries.map((entry) => entry.rule);
  recomputeRuleGridValues(filtered);
  const activeCell = getActiveRuleCellMeta();
  const activeRawValue = activeCell.rule ? getRuleCellRawValue(activeCell.rule, activeCell.field) : "";

  const fixedCount = filtered.filter((rule) => rule.type === "fixed").length;
  const variableCount = filtered.filter((rule) => rule.type === "variable").length;
  els.rulesStats.innerHTML = `
    <div class="rules-entry-toolbar">
      <div class="hc-module-note">当前录入视图支持整列粘贴、公式、Ctrl/Cmd+C 复制选区，以及按筛选结果向下填充</div>
      <div class="hc-formula-bar rules-formula-bar">
        <div class="hc-name-box">${getRuleCellAddress(activeCell.rowIndex, activeCell.colIndex)}</div>
        <label class="hc-formula-field">
          <span>fx</span>
          <input id="rules-formula-input" type="text" value="${activeRawValue}" placeholder="输入数值或公式，例如 =A1*1.1" />
        </label>
      </div>
      <div class="hc-shortcuts rules-shortcuts">
        <button class="hc-action-btn" type="button" data-rules-action="fill-to-end">复制到末行</button>
        <span>Enter 下移</span>
        <span>Shift+Enter 上移</span>
        <span>Tab 横向切换</span>
        <span>Ctrl/Cmd+D 向下填充</span>
      </div>
    </div>
    <article class="rules-stat"><span>规则总数</span><strong>${rules.length}</strong></article>
    <article class="rules-stat"><span>当前显示</span><strong>${filtered.length}</strong></article>
    <article class="rules-stat"><span>固定费用</span><strong>${fixedCount}</strong></article>
    <article class="rules-stat"><span>变动费用</span><strong>${variableCount}</strong></article>
  `;

  els.rulesBody.innerHTML = filteredEntries.length
    ? renderRuleModuleCards(filteredEntries, monthOptions)
    : `<section class="rules-module-card"><div class="empty-state">当前筛选条件下没有匹配的费用科目。</div></section>`;
}

function updateModeUI() {
  const isPost = state.importMode === "post";
  if (els.postUploadWrap) els.postUploadWrap.classList.toggle("hidden", !isPost);
  if (els.splitUploadWrap) els.splitUploadWrap.classList.toggle("hidden", isPost);
  els.scopeFilter.disabled = isPost;
  els.scopeFilter.innerHTML = isPost
    ? `<option value="post">分摊后</option>`
    : `<option value="pre">分摊前</option><option value="allocated">分摊值</option>`;
  els.scopeFilter.value = isPost ? "post" : state.ruleScope;
  const sourceFilter = document.getElementById("source-filter");
  sourceFilter.innerHTML = isPost
    ? `<option value="all">全部</option><option value="分摊后">分摊后</option>`
    : `<option value="all">全部</option><option value="项目组直属">项目组直属</option><option value="分摊值">分摊值</option>`;
  if (![...sourceFilter.options].some((option) => option.value === state.filters.source)) {
    state.filters.source = "all";
  }
  sourceFilter.value = state.filters.source;
  els.rulesSubtitle.textContent = isPost
    ? "当前为分摊后口径：步骤 2.5 - 2.7 三个费用预估模块均按分摊后历史会计科目与金额生成建议值。"
    : `当前为${state.ruleScope === "pre" ? "分摊前" : "分摊值"}口径：步骤 2.5 - 2.7 三个费用预估模块均按对应历史数据生成建议值。`;
}

function renderTabPanels() {
  els.tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.tabTarget === state.activeTab);
  });
  els.tabPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.tabPanel === state.activeTab);
  });
}

function renderChart(model) {
  const maxValue = Math.max(...model.monthly.flatMap((row) => [row.netRevenue, row.merged, Math.abs(row.profit)]), 1);
  const bars = model.monthly.map((row, index) => {
    const x = 50 + index * 68;
    const revenueHeight = (row.netRevenue / maxValue) * 180;
    const costHeight = (row.merged / maxValue) * 180;
    const profitY = 250 - (row.profit / maxValue) * 160;
    return `
      <rect x="${x}" y="${250 - revenueHeight}" width="18" height="${revenueHeight}" fill="#0f6cbd" opacity="0.82"></rect>
      <rect x="${x + 24}" y="${250 - costHeight}" width="18" height="${costHeight}" fill="#f77f00" opacity="0.82"></rect>
      <circle cx="${x + 50}" cy="${profitY}" r="4" fill="${row.profit >= 0 ? "#1f8f6a" : "#c03b33"}"></circle>
      <text x="${x + 18}" y="276" font-size="11" fill="#64748b">${row.month.slice(5)}</text>
    `;
  }).join("");

  els.summaryChart.innerHTML = `
    <line x1="40" y1="250" x2="900" y2="250" stroke="rgba(19,34,56,0.18)"></line>
    <text x="56" y="24" font-size="12" fill="#0f6cbd">净收入</text>
    <text x="112" y="24" font-size="12" fill="#f77f00">合并成本</text>
    <text x="190" y="24" font-size="12" fill="#1f8f6a">项目利润</text>
    ${bars}
  `;
}

function render() {
  const model = calculateModel();
  renderTabPanels();
  updateModeUI();
  renderHistoricalOverview();
  renderStep1PivotPanel();
  renderStep1WhatIf();
  renderPostPivot();
  renderHcModule();
  renderRevenueModule();
  renderCostModule();
  renderMarketingModule();
  renderMetrics(model);
  renderBudgetStrip(model);
  renderRules();
  renderPeriodSummaryTables(model);
  renderProfitBridgeTable(model);
  renderProfitBridgeChart(model);
  renderWaterfallChart(model);
  renderRevenueStructureChart(model);
  renderCostStructureChart(model);
  renderStressTest(model);
  scheduleAutoSave();
}

function syncInputs() {
  els.forecastStart.value = state.forecastStart;
  els.forecastMonths.value = state.forecastMonths;
  els.importMode.value = state.importMode;
  els.ruleSearch.value = state.filters.search;
  els.scopeFilter.value = state.ruleScope;
  updateModeUI();
}

function bindEvents() {
  els.tabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.activeTab = button.dataset.tabTarget;
      renderTabPanels();
      scheduleAutoSave();
    });
  });
  els.draftSlotSelect?.addEventListener("change", (event) => {
    const selectedDraft = getDraftEntryById(event.target.value);
    if (!selectedDraft) return;
    state.persistence.activeDraftId = selectedDraft.id;
    state.persistence.activeDraftName = selectedDraft.name;
    if (els.draftSlotNameInput) {
      els.draftSlotNameInput.value = selectedDraft.name;
    }
    setDraftStatus(`当前已切换到草稿槽位「${selectedDraft.name}」：${formatTimestamp(selectedDraft.savedAt)}`, {
      hasDraft: true,
      lastSavedAt: selectedDraft.savedAt
    });
  });
  els.draftSlotNameInput?.addEventListener("input", (event) => {
    state.persistence.activeDraftName = normalizeDraftName(event.target.value) || getDefaultDraftName();
    scheduleAutoSave();
  });
  els.draftSaveButton?.addEventListener("click", () => {
    persistDraft("manual");
  });
  els.draftSaveAsButton?.addEventListener("click", () => {
    persistDraft("manual", { forceNew: true });
  });
  els.draftRestoreButton?.addEventListener("click", () => {
    const selectedDraft = getDraftEntryById(getSelectedDraftId());
    if (!selectedDraft?.payload) {
      setDraftStatus("没有可恢复的本地草稿槽位", { hasDraft: false, lastSavedAt: "" });
      return;
    }
    if (!window.confirm(`恢复草稿「${selectedDraft.name}」会覆盖当前页面内尚未保存的录入内容，是否继续？`)) return;
    restoreSelectedDraft();
    render();
  });
  els.draftClearButton?.addEventListener("click", () => {
    const selectedDraft = getDraftEntryById(getSelectedDraftId());
    if (!selectedDraft) return;
    if (!window.confirm(`这会删除草稿「${selectedDraft.name}」，但不会删除当前页面已加载的数据。是否继续？`)) return;
    deleteSelectedDraft();
  });
  els.resultActions.forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.exportAction;
      if (action === "summary-csv") exportSummaryCsv();
      if (action === "bridge-csv") exportProfitBridgeCsv();
      if (action === "snapshot-json") exportSnapshotJson();
      if (action === "report-html") exportShareReportHtml();
    });
  });
  els.forecastStart.addEventListener("input", (event) => {
    state.forecastStart = event.target.value;
    rebuildHcForecastRows();
    rebuildRevenueForecastParts();
    rebuildCostForecastRows();
    rebuildMarketingForecastRows();
    syncRuleBoundaries();
    refreshRules();
    render();
  });
  els.forecastMonths.addEventListener("input", (event) => {
    state.forecastMonths = Number(event.target.value || 12);
    rebuildHcForecastRows();
    rebuildRevenueForecastParts();
    rebuildCostForecastRows();
    rebuildMarketingForecastRows();
    render();
  });
  els.importMode.addEventListener("change", (event) => {
    state.importMode = event.target.value;
    state.ruleScope = state.importMode === "post" ? "post" : "pre";
    state.filters.source = "all";
    state.filters.search = "";
    applyBoundaryDefaults();
    updateModeUI();
    render();
  });
  els.scopeFilter.addEventListener("change", (event) => {
    state.ruleScope = event.target.value;
    updateModeUI();
    render();
  });
  document.getElementById("source-filter").addEventListener("change", (event) => { state.filters.source = event.target.value; renderRules(); });
  document.getElementById("rule-filter").addEventListener("change", (event) => { state.filters.type = event.target.value; renderRules(); });
  els.ruleSearch.addEventListener("input", (event) => { state.filters.search = event.target.value; renderRules(); });
  document.getElementById("pnl-post-file-input").addEventListener("change", async (event) => handleFileSelect("post", event.target.files[0]));
  document.getElementById("pnl-pre-file-input").addEventListener("change", async (event) => handleFileSelect("pre", event.target.files[0]));
  document.getElementById("pnl-alloc-file-input").addEventListener("change", async (event) => handleFileSelect("allocated", event.target.files[0]));
  document.getElementById("flow-file-input").addEventListener("change", async (event) => handleFileSelect("flow", event.target.files[0]));
  els.historyBreakdownStartMonth?.addEventListener("change", (event) => {
    state.historicalBreakdownStartMonth = event.target.value;
    if (state.historicalBreakdownMode === "single") {
      state.historicalBreakdownEndMonth = event.target.value;
    }
    renderHistoricalOverview();
    scheduleAutoSave();
  });
  els.historyBreakdownEndMonth?.addEventListener("change", (event) => {
    state.historicalBreakdownEndMonth = event.target.value;
    if (state.historicalBreakdownMode === "single") {
      state.historicalBreakdownStartMonth = event.target.value;
    }
    renderHistoricalOverview();
    scheduleAutoSave();
  });
  els.historyBreakdownModeSwitch?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-breakdown-mode]");
    if (!button) return;
    state.historicalBreakdownMode = button.dataset.breakdownMode;
    if (state.historicalBreakdownMode === "single") {
      state.historicalBreakdownStartMonth = state.historicalBreakdownEndMonth || state.historicalBreakdownStartMonth;
    }
    renderHistoricalOverview();
    scheduleAutoSave();
  });
  els.historyBurnWindow?.addEventListener("change", (event) => {
    state.historicalBurnWindow = Number(event.target.value || 3);
    renderHistoricalOverview();
    renderStep1WhatIf();
    scheduleAutoSave();
  });
  els.historyProjectBudget?.addEventListener("input", (event) => {
    state.projectBudgetAmount = Math.max(parseFormattedNumber(event.target.value || 0), 0);
    renderHistoricalOverview();
    scheduleAutoSave();
  });
  els.historyOverviewSummary?.addEventListener("click", (event) => {
    const toggleButton = event.target.closest("#history-budget-alert-toggle");
    if (!toggleButton) return;
    state.budgetAlertExpanded = !state.budgetAlertExpanded;
    renderHistoricalOverview();
    scheduleAutoSave();
  });
  els.historyOverviewSummary?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.id === "history-budget-alert-launch-month") {
      state.budgetAlertLaunchMonth = normalizeMonth(target.value);
      renderHistoricalOverview();
      scheduleAutoSave();
      return;
    }
    if (target.id === "history-budget-alert-marketing-budget") {
      state.budgetAlertMarketingBudget = Math.max(parseFormattedNumber(target.value || 0), 0);
      renderHistoricalOverview();
      scheduleAutoSave();
      return;
    }
    if (target.id === "history-budget-alert-hc-growth-range") {
      state.budgetAlertHcGrowthRate = Number(target.value || 0) / 100;
      renderHistoricalOverview();
      scheduleAutoSave();
    }
  });
  els.pivotToggleButton?.addEventListener("click", () => {
    state.pivotExpanded = !state.pivotExpanded;
    renderStep1PivotPanel();
    scheduleAutoSave();
  });
  els.delayMonthsRange?.addEventListener("input", (event) => {
    const nextValue = Number(event.target.value);
    state.delayMonths = Number.isFinite(nextValue) ? nextValue : 6;
    renderStep1WhatIf();
    scheduleAutoSave();
  });
  els.delayHcGrowthRange?.addEventListener("input", (event) => {
    state.delayHcGrowthRate = Number(event.target.value || 0) / 100;
    renderStep1WhatIf();
    scheduleAutoSave();
  });
  els.paybackMonthsRange?.addEventListener("input", (event) => {
    state.paybackMonths = Number(event.target.value || 2);
    renderStep1WhatIf();
    scheduleAutoSave();
  });
  els.grossMarginRange?.addEventListener("input", (event) => {
    state.grossMarginRate = Number(event.target.value || 50) / 100;
    renderStep1WhatIf();
    scheduleAutoSave();
  });
  els.revenueViewToggle?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-revenue-view]");
    if (!button) return;
    state.resultRevenueView = button.dataset.revenueView === "platform" ? "platform" : "region";
    renderRevenueStructureChart(calculateModel());
    scheduleAutoSave();
  });
  els.stressRevenueDownRange?.addEventListener("input", (event) => {
    state.stressRevenueDownRate = Number(event.target.value || 0) / 100;
    renderStressTest(calculateModel());
    scheduleAutoSave();
  });
  els.stressUaUpRange?.addEventListener("input", (event) => {
    state.stressUaUpRate = Number(event.target.value || 0) / 100;
    renderStressTest(calculateModel());
    scheduleAutoSave();
  });
  els.rulesBody.addEventListener("input", handleRuleEdit);
  els.rulesBody.addEventListener("change", handleRuleEdit);
  els.rulesStats.addEventListener("click", handleRulesClick);
  els.rulesStats.addEventListener("input", handleRuleEdit);
  els.rulesStats.addEventListener("change", handleRuleEdit);
  els.rulesStats.addEventListener("keydown", handleRulesKeyDown);
  els.rulesBody.addEventListener("focusin", handleRulesFocusIn);
  els.rulesBody.addEventListener("click", handleRulesClick);
  els.rulesBody.addEventListener("dblclick", handleRulesDoubleClick);
  els.rulesBody.addEventListener("copy", handleRulesCopy);
  els.rulesBody.addEventListener("paste", handleRulesPaste);
  els.rulesBody.addEventListener("keydown", handleRulesKeyDown);
  els.rulesBody.addEventListener("mousedown", handleRulesMouseDown);
  els.rulesBody.addEventListener("mouseover", handleRulesMouseOver);
  els.hcModule.addEventListener("focusin", handleHcFocusIn);
  els.hcModule.addEventListener("input", handleHcInput);
  els.hcModule.addEventListener("change", handleHcInput);
  els.hcModule.addEventListener("click", handleHcClick);
  els.hcModule.addEventListener("dblclick", handleHcDoubleClick);
  els.hcModule.addEventListener("copy", handleHcCopy);
  els.hcModule.addEventListener("paste", handleHcPaste);
  els.hcModule.addEventListener("keydown", handleHcKeyDown);
  els.hcModule.addEventListener("mousedown", handleHcMouseDown);
  els.hcModule.addEventListener("mouseover", handleHcMouseOver);
  els.revenueModule.addEventListener("input", handleRevenueInput);
  els.revenueModule.addEventListener("change", handleRevenueInput);
  els.revenueModule.addEventListener("focusin", handleRevenueFocusIn);
  els.revenueModule.addEventListener("click", handleRevenueClick);
  els.revenueModule.addEventListener("dblclick", handleRevenueDoubleClick);
  els.revenueModule.addEventListener("copy", handleRevenueCopy);
  els.revenueModule.addEventListener("paste", handleRevenuePaste);
  els.revenueModule.addEventListener("keydown", handleRevenueKeyDown);
  els.revenueModule.addEventListener("mousedown", handleRevenueMouseDown);
  els.revenueModule.addEventListener("mouseover", handleRevenueMouseOver);
  els.costModule.addEventListener("input", handleCostInput);
  els.costModule.addEventListener("change", handleCostInput);
  els.costModule.addEventListener("focusin", handleCostFocusIn);
  els.costModule.addEventListener("click", handleCostClick);
  els.costModule.addEventListener("dblclick", handleCostDoubleClick);
  els.costModule.addEventListener("copy", handleCostCopy);
  els.costModule.addEventListener("paste", handleCostPaste);
  els.costModule.addEventListener("keydown", handleCostKeyDown);
  els.costModule.addEventListener("mousedown", handleCostMouseDown);
  els.costModule.addEventListener("mouseover", handleCostMouseOver);
  els.marketingModule.addEventListener("input", handleMarketingInput);
  els.marketingModule.addEventListener("change", handleMarketingInput);
  els.marketingModule.addEventListener("focusin", handleMarketingFocusIn);
  els.marketingModule.addEventListener("click", handleMarketingClick);
  els.marketingModule.addEventListener("dblclick", handleMarketingDoubleClick);
  els.marketingModule.addEventListener("copy", handleMarketingCopy);
  els.marketingModule.addEventListener("paste", handleMarketingPaste);
  els.marketingModule.addEventListener("keydown", handleMarketingKeyDown);
  els.marketingModule.addEventListener("mousedown", handleMarketingMouseDown);
  els.marketingModule.addEventListener("mouseover", handleMarketingMouseOver);
  document.addEventListener("mouseup", handleDocumentMouseUp);
}

function handleRuleEdit(event) {
  const formulaInput = event.target.closest("#rules-formula-input");
  if (formulaInput) {
    const activeCell = getActiveRuleCellMeta();
    if (!activeCell.rule || !isRuleEditable(activeCell.rule, activeCell.field)) return;
    const rows = getRenderableRules();
    setRuleCellRawValue(activeCell.rule, activeCell.field, formulaInput.value, { recompute: false, rows });
    if (event.type === "change") {
      recomputeRuleGridValues(rows);
      renderRules();
    }
    return;
  }

  const row = event.target.closest("tr[data-key]");
  if (!row || !event.target.dataset.field) return;
  const rows = getRenderableRules();
  const rule = rows.find((item) => item.key === row.dataset.key);
  if (!rule) return;
  const field = event.target.dataset.field;
  if (["expectedValue", "growthRate"].includes(field)) {
    if (!isRuleEditable(rule, field)) return;
    setRuleCellRawValue(rule, field, event.target.value, { recompute: false, rows });
    state.rulesEditCell = { rowIndex: Number(event.target.dataset.ruleRow), colIndex: Number(event.target.dataset.ruleCol) };
    const formulaBar = document.getElementById("rules-formula-input");
    if (formulaBar) formulaBar.value = event.target.value;
    if (event.type === "change") {
      recomputeRuleGridValues(rows);
      state.rulesEditCell = null;
      render();
    }
    return;
  }

  const value = event.target.value;
  rule[field] = value;
  if (field === "type" && value === "fixed") rule.method = "fixed";
  render();
}

function handleRulesFocusIn(event) {
  const input = event.target.closest("input[data-rule-row][data-rule-col]");
  if (!input) return;
  const rowIndex = Number(input.dataset.ruleRow);
  const colIndex = Number(input.dataset.ruleCol);
  const rule = getRenderableRules()[rowIndex];
  const field = getRuleEditableColumns()[colIndex]?.field;
  if (!rule || !field) return;
  if (!isRuleEditable(rule, field)) {
    input.blur();
    return;
  }
  state.rulesEditCell = { rowIndex, colIndex };
  setRulesSelection(rowIndex, colIndex);
  input.value = getRuleCellRawValue(rule, field);
  input.select();
  const formulaBar = document.getElementById("rules-formula-input");
  if (formulaBar) formulaBar.value = input.value;
}

function handleRulesClick(event) {
  const action = event.target.closest("[data-rules-action]");
  if (action?.dataset.rulesAction === "fill-to-end") {
    fillRulesSelectionToLastRow();
    renderRules();
    focusRuleCellInput(getActiveRuleCell().rowIndex, getActiveRuleCell().colIndex, { selectText: false });
    return;
  }

  const cell = event.target.closest("td[data-rule-row][data-rule-col]");
  if (!cell) return;
  const rowIndex = Number(cell.dataset.ruleRow);
  const colIndex = Number(cell.dataset.ruleCol);
  const rule = getRenderableRules()[rowIndex];
  const field = getRuleEditableColumns()[colIndex]?.field;
  if (!isRuleEditable(rule, field)) return;
  if (!event.target.closest("input")) {
    setActiveRuleCell(rowIndex, colIndex, { selectText: true });
  }
}

function handleRulesCopy(event) {
  if (!event.clipboardData) return;
  event.preventDefault();
  event.clipboardData.setData("text/plain", getRuleSelectionText());
}

function handleRulesPaste(event) {
  const input = event.target.closest("input[data-rule-row][data-rule-col]");
  if (!input) return;
  const text = event.clipboardData?.getData("text");
  if (!text) return;

  event.preventDefault();
  const rows = getRenderableRules();
  const columns = getRuleEditableColumns();
  const startRowIndex = Number(input.dataset.ruleRow);
  const startColIndex = Number(input.dataset.ruleCol);
  let appliedEndRowIndex = startRowIndex;
  let appliedEndColIndex = startColIndex;

  text
    .trim()
    .split(/\r?\n/)
    .forEach((line, rowOffset) => {
      const rule = rows[startRowIndex + rowOffset];
      if (!rule) return;
      appliedEndRowIndex = startRowIndex + rowOffset;
      line.split("\t").forEach((cell, colOffset) => {
        const column = columns[startColIndex + colOffset];
        if (!column) return;
        appliedEndColIndex = Math.max(appliedEndColIndex, startColIndex + colOffset);
        setRuleCellRawValue(rule, column.field, cell, { recompute: false, rows });
      });
    });

  recomputeRuleGridValues(rows);
  setRulesSelection(startRowIndex, startColIndex, appliedEndRowIndex, appliedEndColIndex);
  state.rulesEditCell = { rowIndex: startRowIndex, colIndex: startColIndex };
  renderRules();
}

function handleRulesMouseDown(event) {
  const handle = event.target.closest("[data-rules-fill-handle]");
  if (!handle) return;
  event.preventDefault();
  const selection = normalizeRulesSelection();
  state.rulesFillDrag = { targetRow: selection.endRow };
}

function handleRulesMouseOver(event) {
  if (!state.rulesFillDrag) return;
  const cell = event.target.closest("td[data-rule-row][data-rule-col]");
  if (!cell) return;
  const rowIndex = Number(cell.dataset.ruleRow);
  state.rulesFillDrag.targetRow = Math.max(rowIndex, normalizeRulesSelection().endRow);
  renderRules();
}

function handleRulesDoubleClick(event) {
  const handle = event.target.closest("[data-rules-fill-handle]");
  if (!handle) return;
  event.preventDefault();
  fillRulesSelectionToLastRow();
  renderRules();
  focusRuleCellInput(getActiveRuleCell().rowIndex, getActiveRuleCell().colIndex, { selectText: false });
}

function handleRulesKeyDown(event) {
  const formulaInput = event.target.closest("#rules-formula-input");
  if (formulaInput && event.key === "Enter") {
    event.preventDefault();
    const activeCell = getActiveRuleCellMeta();
    if (!activeCell.rule) return;
    const rows = getRenderableRules();
    setRuleCellRawValue(activeCell.rule, activeCell.field, formulaInput.value, { recompute: false, rows });
    recomputeRuleGridValues(rows);
    renderRules();
    focusRuleCellInput(activeCell.rowIndex, activeCell.colIndex);
    return;
  }

  const input = event.target.closest("input[data-rule-row][data-rule-col]");
  if (!input) return;
  const rows = getRenderableRules();

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    const rule = rows[Number(input.dataset.ruleRow)];
    const field = getRuleEditableColumns()[Number(input.dataset.ruleCol)]?.field;
    setRuleCellRawValue(rule, field, input.value, { recompute: false, rows });
    recomputeRuleGridValues(rows);
    fillRulesSelectionToLastRow();
    renderRules();
    focusRuleCellInput(Number(input.dataset.ruleRow), Number(input.dataset.ruleCol), { selectText: false });
    return;
  }

  let nextRowIndex = Number(input.dataset.ruleRow);
  let nextColIndex = Number(input.dataset.ruleCol);
  if (event.key === "Enter") nextRowIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "Tab") nextColIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "ArrowUp") nextRowIndex -= 1;
  else if (event.key === "ArrowDown") nextRowIndex += 1;
  else if (event.key === "ArrowLeft") {
    if ((input.selectionStart ?? 0) !== 0 || (input.selectionEnd ?? 0) !== 0) return;
    nextColIndex -= 1;
  } else if (event.key === "ArrowRight") {
    const caretEnd = input.value.length;
    if ((input.selectionStart ?? caretEnd) !== caretEnd || (input.selectionEnd ?? caretEnd) !== caretEnd) return;
    nextColIndex += 1;
  } else {
    return;
  }

  event.preventDefault();
  const rule = rows[Number(input.dataset.ruleRow)];
  const field = getRuleEditableColumns()[Number(input.dataset.ruleCol)]?.field;
  setRuleCellRawValue(rule, field, input.value, { recompute: false, rows });
  recomputeRuleGridValues(rows);
  setActiveRuleCell(nextRowIndex, nextColIndex);
}

function updateHcValue(code, month, value, options) {
  setHcCellRawValue(code, month, value, options);
}

function handleHcInput(event) {
  const metaInput = event.target.closest("[data-hc-meta-field][data-hc-meta-code]");
  if (metaInput) {
    const row = getHcRowByCode(metaInput.dataset.hcMetaCode);
    if (!row) return;
    if (metaInput.dataset.hcMetaField === "feeType") return;
    row[metaInput.dataset.hcMetaField] = metaInput.value;
    if (metaInput.dataset.hcMetaField === "estimateMethodRaw") {
      row.forecastLogicRaw = deriveHcForecastLogic(row.estimateMethodRaw, row.forecastLogicRaw);
    }
    if (metaInput.dataset.hcMetaField === "referencePeriodRaw") {
      row.suggestedValueText = deriveHcSuggestedValueText(row.code, row.referencePeriodRaw);
    }
    if (row.computed && metaInput.dataset.hcMetaField === "estimateMethodRaw") {
      getHcMonths().forEach((month) => {
        setHcCellRawValue(row.code, month, metaInput.value, { recompute: false });
      });
      recomputeHcForecastValues();
    }
    if (event.type === "change" || ["referencePeriodRaw", "estimateMethodRaw", "forecastParamRaw"].includes(metaInput.dataset.hcMetaField)) {
      renderHcModule();
      if (event.type === "input" && metaInput.dataset.hcMetaField === "forecastParamRaw") {
        refocusMetaFieldInput(els.hcModule, `[data-hc-meta-field="forecastParamRaw"][data-hc-meta-code="${row.code}"]`);
      }
    }
    return;
  }

  const formulaInput = event.target.closest("#hc-formula-input");
  if (formulaInput) {
    const activeCell = getActiveHcCellMeta();
    if (getHcRowByCode(activeCell.code)?.computed) return;
    updateHcValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    if (event.type === "change") {
      recomputeHcForecastValues();
      render();
    }
    return;
  }
  const input = event.target.closest("input[data-hc-code][data-hc-month]");
  if (!input) return;
  if (getHcRowByCode(input.dataset.hcCode)?.computed) return;
  updateHcValue(input.dataset.hcCode, input.dataset.hcMonth, input.value, { recompute: false });
  state.hcEditCell = { rowIndex: Number(input.dataset.hcRow), colIndex: Number(input.dataset.hcCol) };
  const formulaBar = document.getElementById("hc-formula-input");
  if (formulaBar) formulaBar.value = input.value;
  if (event.type === "change") {
    recomputeHcForecastValues();
    state.hcEditCell = null;
    render();
  }
}

function handleHcPaste(event) {
  const text = event.clipboardData?.getData("text");
  if (!text) return;

  event.preventDefault();
  const months = getHcMonths();
  const activeCell = getActiveHcCellMeta();
  if (getHcRowByCode(activeCell.code)?.computed) return;
  const startRowIndex = activeCell.rowIndex;
  const startColIndex = activeCell.colIndex;
  if (startRowIndex < 0 || startColIndex < 0) return;

  let appliedEndRowIndex = startRowIndex;
  let appliedEndColIndex = startColIndex;

  text
    .trim()
    .split(/\r?\n/)
    .forEach((line, rowOffset) => {
      const code = hcAccountCodes[startRowIndex + rowOffset];
      if (!code) return;
      appliedEndRowIndex = startRowIndex + rowOffset;
      line.split("\t").forEach((cell, colOffset) => {
        const month = months[startColIndex + colOffset];
        if (!month) return;
        appliedEndColIndex = Math.max(appliedEndColIndex, startColIndex + colOffset);
        updateHcValue(code, month, cell, { recompute: false });
      });
    });

  recomputeHcForecastValues();
  setHcSelection(startRowIndex, startColIndex, appliedEndRowIndex, appliedEndColIndex);
  state.hcEditCell = { rowIndex: startRowIndex, colIndex: startColIndex };
  render();
}

function handleHcFocusIn(event) {
  const input = event.target.closest("input[data-hc-code][data-hc-month]");
  if (!input) return;
  const rowIndex = Number(input.dataset.hcRow);
  const colIndex = Number(input.dataset.hcCol);
  state.hcEditCell = { rowIndex, colIndex };
  setHcSelection(rowIndex, colIndex);
  input.value = getHcCellRawValue(input.dataset.hcCode, input.dataset.hcMonth);
  input.select();
  const formulaBar = document.getElementById("hc-formula-input");
  if (formulaBar) {
    formulaBar.value = input.value;
    formulaBar.readOnly = Boolean(getHcRowByCode(input.dataset.hcCode)?.computed);
  }
}

function handleHcClick(event) {
  if (state.hcFormulaInsertLock) {
    state.hcFormulaInsertLock = false;
    return;
  }

  const fillHandle = event.target.closest("[data-hc-fill-handle]");
  if (fillHandle) {
    const cell = fillHandle.closest("td[data-hc-row][data-hc-col]");
    if (cell && event.detail >= 2) {
      event.preventDefault();
      event.stopPropagation();
      performHcFillToLastMonth(Number(cell.dataset.hcRow), Number(cell.dataset.hcCol));
      return;
    }
  }

  const actionButton = event.target.closest("[data-hc-action]");
  if (actionButton?.dataset.hcAction === "fill-to-end") {
    fillHcSelectionToLastRow();
    render();
    focusHcCellInput(getActiveHcCell().rowIndex, getActiveHcCell().colIndex, { selectText: false });
    return;
  }

  const header = event.target.closest("[data-hc-row-header]");
  if (header) {
    const rowIndex = Number(header.dataset.hcRowHeader);
    setHcSelection(rowIndex, 0, rowIndex, getHcMonths().length - 1, "row");
    state.hcEditCell = null;
    renderHcModule();
    return;
  }

  const cell = event.target.closest("td[data-hc-row][data-hc-col]");
  if (!cell) return;
  const rowIndex = Number(cell.dataset.hcRow);
  const colIndex = Number(cell.dataset.hcCol);
  setActiveHcCell(rowIndex, colIndex, { selectText: !event.target.closest("input") });
}

function handleHcCopy(event) {
  if (!event.clipboardData) return;
  if (shouldUseNativeCopy(event)) return;
  event.preventDefault();
  event.clipboardData.setData("text/plain", getHcSelectionText());
}

function handleHcMouseDown(event) {
  const handle = event.target.closest("[data-hc-fill-handle]");
  if (handle) {
    event.preventDefault();
    event.stopPropagation();
    const cell = handle.closest("td[data-hc-row][data-hc-col]");
    if (cell) {
      const rowIndex = Number(cell.dataset.hcRow);
      const colIndex = Number(cell.dataset.hcCol);
      setHcSelection(rowIndex, colIndex);
      state.hcEditCell = { rowIndex, colIndex };
    }
    const selection = normalizeHcSelection();
    state.hcFillDrag = { targetCol: selection.endCol, startCol: selection.endCol, dragging: false };
    return;
  }

  const cell = event.target.closest("td[data-hc-row][data-hc-col]");
  if (!cell) return;
  if (tryInsertHcFormulaReferenceFromCell(cell)) {
    event.preventDefault();
  }
}

function handleHcMouseOver(event) {
  if (!state.hcFillDrag) return;
  const cell = event.target.closest("td[data-hc-row][data-hc-col]");
  if (!cell) return;
  const colIndex = Number(cell.dataset.hcCol);
  state.hcFillDrag.targetCol = Math.max(colIndex, normalizeHcSelection().endCol);
  state.hcFillDrag.dragging = state.hcFillDrag.targetCol > state.hcFillDrag.startCol;
  renderHcModule();
}

function handleHcDoubleClick(event) {
  const handle = event.target.closest("[data-hc-fill-handle]");
  if (!handle) return;
  event.preventDefault();
  event.stopPropagation();
  const cell = handle.closest("td[data-hc-row][data-hc-col]");
  if (!cell) return;
  performHcFillToLastMonth(Number(cell.dataset.hcRow), Number(cell.dataset.hcCol));
}

function handleHcKeyDown(event) {
  const formulaInput = event.target.closest("#hc-formula-input");
  if (formulaInput && event.key === "Enter") {
    event.preventDefault();
    const activeCell = getActiveHcCellMeta();
    updateHcValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    recomputeHcForecastValues();
    renderHcModule();
    focusHcCellInput(activeCell.rowIndex, activeCell.colIndex);
    return;
  }

  const input = event.target.closest("input[data-hc-code][data-hc-month]");
  if (!input) return;
  if (getHcRowByCode(input.dataset.hcCode)?.computed) return;

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    updateHcValue(input.dataset.hcCode, input.dataset.hcMonth, input.value, { recompute: false });
    recomputeHcForecastValues();
    fillHcSelectionToLastRow();
    render();
    focusHcCellInput(Number(input.dataset.hcRow), Number(input.dataset.hcCol), { selectText: false });
    return;
  }

  let nextRowIndex = Number(input.dataset.hcRow);
  let nextColIndex = Number(input.dataset.hcCol);

  if (event.key === "Enter") nextRowIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "Tab") nextColIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "ArrowUp") nextRowIndex -= 1;
  else if (event.key === "ArrowDown") nextRowIndex += 1;
  else if (event.key === "ArrowLeft") {
    if ((input.selectionStart ?? 0) !== 0 || (input.selectionEnd ?? 0) !== 0) return;
    nextColIndex -= 1;
  } else if (event.key === "ArrowRight") {
    const caretEnd = input.value.length;
    if ((input.selectionStart ?? caretEnd) !== caretEnd || (input.selectionEnd ?? caretEnd) !== caretEnd) return;
    nextColIndex += 1;
  } else {
    return;
  }

  event.preventDefault();
  updateHcValue(input.dataset.hcCode, input.dataset.hcMonth, input.value, { recompute: false });
  recomputeHcForecastValues();
  setActiveHcCell(nextRowIndex, nextColIndex);
}

function handleDocumentMouseUp() {
  if (state.hcFillDrag) {
    if (state.hcFillDrag.dragging && state.hcFillDrag.targetCol > normalizeHcSelection().endCol) {
      fillHcSelectionRight(state.hcFillDrag.targetCol);
      state.hcFillDrag = null;
      render();
      return;
    }
    state.hcFillDrag = null;
    return;
  }
  if (state.revenueFillDrag) {
    if (state.revenueFillDrag.dragging && state.revenueFillDrag.targetCol > normalizeRevenueSelection().endCol) {
      fillRevenueSelectionRight(state.revenueFillDrag.targetCol);
      state.revenueFillDrag = null;
      render();
      return;
    }
    state.revenueFillDrag = null;
    return;
  }
  if (state.costFillDrag) {
    if (state.costFillDrag.dragging && state.costFillDrag.targetCol > normalizeCostSelection().endCol) {
      fillCostSelectionRight(state.costFillDrag.targetCol);
      state.costFillDrag = null;
      render();
      return;
    }
    state.costFillDrag = null;
    return;
  }
  if (state.marketingFillDrag) {
    if (state.marketingFillDrag.dragging && state.marketingFillDrag.targetCol > normalizeMarketingSelection().endCol) {
      fillMarketingSelectionRight(state.marketingFillDrag.targetCol);
      state.marketingFillDrag = null;
      render();
      return;
    }
    state.marketingFillDrag = null;
    return;
  }
  if (state.rulesFillDrag) {
    fillRulesSelectionDown(state.rulesFillDrag.targetRow);
    state.rulesFillDrag = null;
    render();
  }
}

function updateRevenueValue(code, month, value, options) {
  setRevenueCellRawValue(code, month, value, options);
}

function handleRevenueInput(event) {
  const settingInput = event.target.closest("[data-revenue-setting]");
  if (settingInput?.dataset.revenueSetting === "goLiveDate") {
    state.goLiveDate = settingInput.value;
    if (event.type === "change") render();
    return;
  }

  const metaInput = event.target.closest("[data-revenue-meta-field][data-revenue-meta-code]");
  if (metaInput) {
    const row = getRevenueRowByCode(metaInput.dataset.revenueMetaCode);
    if (!row) return;
    row[metaInput.dataset.revenueMetaField] = metaInput.value;
    if (metaInput.dataset.revenueMetaField === "estimateMethodRaw") {
      row.forecastLogicRaw = deriveRevenueForecastLogic(row.estimateMethodRaw, row.forecastLogicRaw);
      if (row.computed) {
        getRevenueMonths().forEach((month) => {
          row.rawValues[month] = row.estimateMethodRaw;
        });
      } else if (isRateEstimateMethod(row.estimateMethodRaw)) {
        getRevenueMonths().forEach((month) => {
          row.rawValues[month] = "";
        });
      }
    }
    if (metaInput.dataset.revenueMetaField === "referencePeriodRaw" || metaInput.dataset.revenueMetaField === "costDriver" || metaInput.dataset.revenueMetaField === "estimateMethodRaw") {
      row.suggestedValueText = deriveRevenueSuggestedValueText(row.code, row.referencePeriodRaw);
      if (row.estimateMethodRaw === "费用率" && metaInput.dataset.revenueMetaField !== "forecastParamRaw" && row.suggestedValueText !== "-") {
        row.forecastParamRaw = row.forecastParamRaw || row.suggestedValueText;
      }
    }
    recomputeRevenueForecastValues();
    if (event.type === "change" || metaInput.dataset.revenueMetaField !== "forecastParamRaw") {
      renderRevenueModule();
      if (event.type === "input" && metaInput.dataset.revenueMetaField === "forecastParamRaw") {
        refocusMetaFieldInput(els.revenueModule, `[data-revenue-meta-field="forecastParamRaw"][data-revenue-meta-code="${row.code}"]`);
      }
    } else if (metaInput.dataset.revenueMetaField === "forecastParamRaw") {
      renderRevenueModule();
      refocusMetaFieldInput(els.revenueModule, `[data-revenue-meta-field="forecastParamRaw"][data-revenue-meta-code="${row.code}"]`);
    }
    return;
  }

  const formulaInput = event.target.closest("#revenue-formula-input");
  if (formulaInput) {
    const activeCell = getActiveRevenueCellMeta();
    updateRevenueValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    if (event.type === "change") {
      recomputeRevenueForecastValues();
      render();
    }
    return;
  }
  const input = event.target.closest("input[data-revenue-code][data-revenue-month]");
  if (!input) return;
  updateRevenueValue(input.dataset.revenueCode, input.dataset.revenueMonth, input.value, { recompute: false });
  state.revenueEditCell = { rowIndex: Number(input.dataset.revenueRow), colIndex: Number(input.dataset.revenueCol) };
  const formulaBar = document.getElementById("revenue-formula-input");
  if (formulaBar) formulaBar.value = input.value;
  if (event.type === "change") {
    recomputeRevenueForecastValues();
    state.revenueEditCell = null;
    render();
  }
}

function handleRevenueClick(event) {
  if (state.revenueFormulaInsertLock) {
    state.revenueFormulaInsertLock = false;
    return;
  }

  const fillHandle = event.target.closest("[data-revenue-fill-handle]");
  if (fillHandle) {
    const cell = fillHandle.closest("td[data-revenue-row][data-revenue-col]");
    if (cell && event.detail >= 2) {
      event.preventDefault();
      event.stopPropagation();
      setRevenueSelection(Number(cell.dataset.revenueRow), Number(cell.dataset.revenueCol));
      state.revenueEditCell = { rowIndex: Number(cell.dataset.revenueRow), colIndex: Number(cell.dataset.revenueCol) };
      fillRevenueSelectionToLastRow();
      render();
      focusRevenueCellInput(Number(cell.dataset.revenueRow), Number(cell.dataset.revenueCol), { selectText: false });
      return;
    }
  }

  const action = event.target.closest("[data-revenue-action]");
  if (action?.dataset.revenueAction === "fill-to-end") {
    fillRevenueSelectionToLastRow();
    render();
    focusRevenueCellInput(getActiveRevenueCell().rowIndex, getActiveRevenueCell().colIndex, { selectText: false });
    return;
  }
  if (action?.dataset.revenueAction === "restore-imported-flow") {
    rebuildRevenueForecastRows({ resetRevenueFromFlow: true });
    recomputeRevenueForecastValues();
    render();
    return;
  }

  const header = event.target.closest("[data-revenue-row-header]");
  if (header) {
    const rowIndex = Number(header.dataset.revenueRowHeader);
    setRevenueSelection(rowIndex, 0, rowIndex, getRevenueMonths().length - 1, "row");
    state.revenueEditCell = null;
    renderRevenueModule();
    return;
  }

  const cell = event.target.closest("td[data-revenue-row][data-revenue-col]");
  if (!cell) return;
  if (tryInsertRevenueFormulaReferenceFromCell(cell)) {
    event.preventDefault();
    return;
  }
  const rowIndex = Number(cell.dataset.revenueRow);
  const colIndex = Number(cell.dataset.revenueCol);
  if (!event.target.closest("input")) {
    setActiveRevenueCell(rowIndex, colIndex, { selectText: true });
  }
}

function handleRevenuePaste(event) {
  const input = event.target.closest("input[data-revenue-code][data-revenue-month]");
  if (!input) return;

  const text = event.clipboardData?.getData("text");
  if (!text) return;

  event.preventDefault();
  const months = getRevenueMonths();
  const startRowIndex = revenueAccountCodes.indexOf(input.dataset.revenueCode);
  const startColIndex = months.indexOf(input.dataset.revenueMonth);
  if (startRowIndex < 0 || startColIndex < 0) return;

  let appliedEndRowIndex = startRowIndex;
  let appliedEndColIndex = startColIndex;
  text
    .trim()
    .split(/\r?\n/)
    .forEach((line, rowOffset) => {
      const code = revenueAccountCodes[startRowIndex + rowOffset];
      if (!code) return;
      appliedEndRowIndex = startRowIndex + rowOffset;
      line.split("\t").forEach((cell, colOffset) => {
        const month = months[startColIndex + colOffset];
        if (!month) return;
        appliedEndColIndex = Math.max(appliedEndColIndex, startColIndex + colOffset);
        updateRevenueValue(code, month, cell, { recompute: false });
      });
    });

  recomputeRevenueForecastValues();
  setRevenueSelection(startRowIndex, startColIndex, appliedEndRowIndex, appliedEndColIndex);
  state.revenueEditCell = { rowIndex: startRowIndex, colIndex: startColIndex };
  render();
}

function handleRevenueFocusIn(event) {
  const input = event.target.closest("input[data-revenue-code][data-revenue-month]");
  if (!input) return;
  const rowIndex = Number(input.dataset.revenueRow);
  const colIndex = Number(input.dataset.revenueCol);
  state.revenueEditCell = { rowIndex, colIndex };
  setRevenueSelection(rowIndex, colIndex);
  input.value = getRevenueCellEditValue(input.dataset.revenueCode, input.dataset.revenueMonth);
  input.select();
  const formulaBar = document.getElementById("revenue-formula-input");
  if (formulaBar) {
    formulaBar.value = input.value;
    formulaBar.readOnly = isAutoCalculatedForecastRow(getRevenueRowByCode(input.dataset.revenueCode));
  }
}

function handleRevenueCopy(event) {
  if (!event.clipboardData) return;
  if (shouldUseNativeCopy(event)) return;
  event.preventDefault();
  event.clipboardData.setData("text/plain", getRevenueSelectionText());
}

function handleRevenueMouseDown(event) {
  const handle = event.target.closest("[data-revenue-fill-handle]");
  if (handle) {
    event.preventDefault();
    event.stopPropagation();
    const cell = handle.closest("td[data-revenue-row][data-revenue-col]");
    if (cell) {
      const rowIndex = Number(cell.dataset.revenueRow);
      const colIndex = Number(cell.dataset.revenueCol);
      setRevenueSelection(rowIndex, colIndex);
      state.revenueEditCell = { rowIndex, colIndex };
    }
    const selection = normalizeRevenueSelection();
    state.revenueFillDrag = { targetCol: selection.endCol, startCol: selection.endCol, dragging: false };
    return;
  }

  const cell = event.target.closest("td[data-revenue-row][data-revenue-col]");
  if (!cell) return;
  if (event.target.closest("input")) return;
  const rowIndex = Number(cell.dataset.revenueRow);
  const colIndex = Number(cell.dataset.revenueCol);
  setRevenueSelection(rowIndex, colIndex);
  state.revenueEditCell = null;
  renderRevenueModule();
}

function handleRevenueMouseOver(event) {
  if (!state.revenueFillDrag) return;
  const cell = event.target.closest("td[data-revenue-row][data-revenue-col]");
  if (!cell) return;
  const colIndex = Number(cell.dataset.revenueCol);
  if (colIndex > normalizeRevenueSelection().endCol) {
    state.revenueFillDrag.dragging = true;
  }
  state.revenueFillDrag.targetCol = Math.max(colIndex, normalizeRevenueSelection().endCol);
  renderRevenueModule();
}

function handleRevenueDoubleClick(event) {
  const handle = event.target.closest("[data-revenue-fill-handle]");
  if (!handle) return;
  event.preventDefault();
  fillRevenueSelectionToLastRow();
  render();
  focusRevenueCellInput(getActiveRevenueCell().rowIndex, getActiveRevenueCell().colIndex, { selectText: false });
}

function handleRevenueKeyDown(event) {
  const formulaInput = event.target.closest("#revenue-formula-input");
  if (formulaInput && event.key === "Enter") {
    event.preventDefault();
    const activeCell = getActiveRevenueCellMeta();
    updateRevenueValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    recomputeRevenueForecastValues();
    renderRevenueModule();
    focusRevenueCellInput(activeCell.rowIndex, activeCell.colIndex);
    return;
  }

  const input = event.target.closest("input[data-revenue-code][data-revenue-month]");
  if (!input) return;

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    updateRevenueValue(input.dataset.revenueCode, input.dataset.revenueMonth, input.value, { recompute: false });
    recomputeRevenueForecastValues();
    fillRevenueSelectionToLastRow();
    render();
    focusRevenueCellInput(Number(input.dataset.revenueRow), Number(input.dataset.revenueCol), { selectText: false });
    return;
  }

  let nextRowIndex = Number(input.dataset.revenueRow);
  let nextColIndex = Number(input.dataset.revenueCol);

  if (event.key === "Enter") nextRowIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "Tab") nextColIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "ArrowUp") nextRowIndex -= 1;
  else if (event.key === "ArrowDown") nextRowIndex += 1;
  else if (event.key === "ArrowLeft") {
    if ((input.selectionStart ?? 0) !== 0 || (input.selectionEnd ?? 0) !== 0) return;
    nextColIndex -= 1;
  } else if (event.key === "ArrowRight") {
    const caretEnd = input.value.length;
    if ((input.selectionStart ?? caretEnd) !== caretEnd || (input.selectionEnd ?? caretEnd) !== caretEnd) return;
    nextColIndex += 1;
  } else {
    return;
  }

  event.preventDefault();
  updateRevenueValue(input.dataset.revenueCode, input.dataset.revenueMonth, input.value, { recompute: false });
  recomputeRevenueForecastValues();
  setActiveRevenueCell(nextRowIndex, nextColIndex);
}

function updateCostValue(code, month, value, options) {
  setCostCellRawValue(code, month, value, options);
}

function handleCostInput(event) {
  const metaInput = event.target.closest("[data-cost-meta-field][data-cost-meta-code]");
  if (metaInput) {
    const row = getCostRowByCode(metaInput.dataset.costMetaCode);
    if (!row) return;
    row[metaInput.dataset.costMetaField] = metaInput.value;
    if (metaInput.dataset.costMetaField === "estimateMethodRaw") {
      row.forecastLogicRaw = deriveCostForecastLogic(row.estimateMethodRaw, row.forecastLogicRaw);
      if (row.computed) {
        getCostMonths().forEach((month) => { row.rawValues[month] = row.estimateMethodRaw; });
      } else if (isRateEstimateMethod(row.estimateMethodRaw)) {
        getCostMonths().forEach((month) => { row.rawValues[month] = ""; });
      }
    }
    if (["referencePeriodRaw", "costDriver", "estimateMethodRaw"].includes(metaInput.dataset.costMetaField)) {
      row.suggestedValueText = deriveCostSuggestedValueText(row.code, row.referencePeriodRaw);
      if (row.estimateMethodRaw === "费用率" && metaInput.dataset.costMetaField !== "forecastParamRaw" && row.suggestedValueText !== "-") {
        row.forecastParamRaw = row.forecastParamRaw || row.suggestedValueText;
      }
    }
    recomputeCostForecastValues();
    if (event.type === "change" || metaInput.dataset.costMetaField !== "forecastParamRaw") {
      renderCostModule();
      if (event.type === "input" && metaInput.dataset.costMetaField === "forecastParamRaw") {
        refocusMetaFieldInput(els.costModule, `[data-cost-meta-field="forecastParamRaw"][data-cost-meta-code="${row.code}"]`);
      }
    } else if (metaInput.dataset.costMetaField === "forecastParamRaw") {
      renderCostModule();
      refocusMetaFieldInput(els.costModule, `[data-cost-meta-field="forecastParamRaw"][data-cost-meta-code="${row.code}"]`);
    }
    return;
  }
  const formulaInput = event.target.closest("#cost-formula-input");
  if (formulaInput) {
    const activeCell = getActiveCostCellMeta();
    updateCostValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    if (event.type === "change") {
      recomputeCostForecastValues();
      render();
    }
    return;
  }
  const input = event.target.closest("input[data-cost-code][data-cost-month]");
  if (!input) return;
  updateCostValue(input.dataset.costCode, input.dataset.costMonth, input.value, { recompute: false });
  state.costEditCell = { rowIndex: Number(input.dataset.costRow), colIndex: Number(input.dataset.costCol) };
  const formulaBar = document.getElementById("cost-formula-input");
  if (formulaBar) formulaBar.value = input.value;
  if (event.type === "change") {
    recomputeCostForecastValues();
    state.costEditCell = null;
    render();
  }
}

function handleCostClick(event) {
  if (state.costFormulaInsertLock) {
    state.costFormulaInsertLock = false;
    return;
  }
  const fillHandle = event.target.closest("[data-cost-fill-handle]");
  if (fillHandle) {
    const cell = fillHandle.closest("td[data-cost-row][data-cost-col]");
    if (cell && event.detail >= 2) {
      event.preventDefault();
      event.stopPropagation();
      setCostSelection(Number(cell.dataset.costRow), Number(cell.dataset.costCol));
      state.costEditCell = { rowIndex: Number(cell.dataset.costRow), colIndex: Number(cell.dataset.costCol) };
      fillCostSelectionToLastRow();
      render();
      focusCostCellInput(Number(cell.dataset.costRow), Number(cell.dataset.costCol), { selectText: false });
      return;
    }
  }
  const action = event.target.closest("[data-cost-action]");
  if (action?.dataset.costAction === "fill-to-end") {
    fillCostSelectionToLastRow();
    render();
    focusCostCellInput(getActiveCostCell().rowIndex, getActiveCostCell().colIndex, { selectText: false });
    return;
  }
  const header = event.target.closest("[data-cost-row-header]");
  if (header) {
    const rowIndex = Number(header.dataset.costRowHeader);
    setCostSelection(rowIndex, 0, rowIndex, getCostMonths().length - 1, "row");
    state.costEditCell = null;
    renderCostModule();
    return;
  }
  const cell = event.target.closest("td[data-cost-row][data-cost-col]");
  if (!cell) return;
  if (tryInsertCostFormulaReferenceFromCell(cell)) {
    event.preventDefault();
    return;
  }
  const rowIndex = Number(cell.dataset.costRow);
  const colIndex = Number(cell.dataset.costCol);
  if (!event.target.closest("input")) setActiveCostCell(rowIndex, colIndex, { selectText: true });
}

function handleCostPaste(event) {
  const input = event.target.closest("input[data-cost-code][data-cost-month]");
  if (!input) return;
  const text = event.clipboardData?.getData("text");
  if (!text) return;
  event.preventDefault();
  const months = getCostMonths();
  const startRowIndex = costAccountCodes.indexOf(input.dataset.costCode);
  const startColIndex = months.indexOf(input.dataset.costMonth);
  if (startRowIndex < 0 || startColIndex < 0) return;
  let appliedEndRowIndex = startRowIndex;
  let appliedEndColIndex = startColIndex;
  text.trim().split(/\r?\n/).forEach((line, rowOffset) => {
    const code = costAccountCodes[startRowIndex + rowOffset];
    if (!code) return;
    appliedEndRowIndex = startRowIndex + rowOffset;
    line.split("\t").forEach((cell, colOffset) => {
      const month = months[startColIndex + colOffset];
      if (!month) return;
      appliedEndColIndex = Math.max(appliedEndColIndex, startColIndex + colOffset);
      updateCostValue(code, month, cell, { recompute: false });
    });
  });
  recomputeCostForecastValues();
  setCostSelection(startRowIndex, startColIndex, appliedEndRowIndex, appliedEndColIndex);
  state.costEditCell = { rowIndex: startRowIndex, colIndex: startColIndex };
  render();
}

function handleCostFocusIn(event) {
  const input = event.target.closest("input[data-cost-code][data-cost-month]");
  if (!input) return;
  const rowIndex = Number(input.dataset.costRow);
  const colIndex = Number(input.dataset.costCol);
  state.costEditCell = { rowIndex, colIndex };
  setCostSelection(rowIndex, colIndex);
  input.value = getCostCellEditValue(input.dataset.costCode, input.dataset.costMonth);
  input.select();
  const formulaBar = document.getElementById("cost-formula-input");
  if (formulaBar) {
    formulaBar.value = input.value;
    formulaBar.readOnly = isAutoCalculatedForecastRow(getCostRowByCode(input.dataset.costCode));
  }
}

function handleCostCopy(event) {
  if (!event.clipboardData) return;
  if (shouldUseNativeCopy(event)) return;
  event.preventDefault();
  event.clipboardData.setData("text/plain", getCostSelectionText());
}

function handleCostMouseDown(event) {
  const handle = event.target.closest("[data-cost-fill-handle]");
  if (handle) {
    event.preventDefault();
    event.stopPropagation();
    const cell = handle.closest("td[data-cost-row][data-cost-col]");
    if (cell) {
      const rowIndex = Number(cell.dataset.costRow);
      const colIndex = Number(cell.dataset.costCol);
      setCostSelection(rowIndex, colIndex);
      state.costEditCell = { rowIndex, colIndex };
    }
    const selection = normalizeCostSelection();
    state.costFillDrag = { targetCol: selection.endCol, startCol: selection.endCol, dragging: false };
    return;
  }
  const cell = event.target.closest("td[data-cost-row][data-cost-col]");
  if (!cell) return;
  if (event.target.closest("input")) return;
  const rowIndex = Number(cell.dataset.costRow);
  const colIndex = Number(cell.dataset.costCol);
  setCostSelection(rowIndex, colIndex);
  state.costEditCell = null;
  renderCostModule();
}

function handleCostMouseOver(event) {
  if (!state.costFillDrag) return;
  const cell = event.target.closest("td[data-cost-row][data-cost-col]");
  if (!cell) return;
  const colIndex = Number(cell.dataset.costCol);
  if (colIndex > normalizeCostSelection().endCol) state.costFillDrag.dragging = true;
  state.costFillDrag.targetCol = Math.max(colIndex, normalizeCostSelection().endCol);
  renderCostModule();
}

function handleCostDoubleClick(event) {
  const handle = event.target.closest("[data-cost-fill-handle]");
  if (!handle) return;
  event.preventDefault();
  fillCostSelectionToLastRow();
  render();
  focusCostCellInput(getActiveCostCell().rowIndex, getActiveCostCell().colIndex, { selectText: false });
}

function handleCostKeyDown(event) {
  const formulaInput = event.target.closest("#cost-formula-input");
  if (formulaInput && event.key === "Enter") {
    event.preventDefault();
    const activeCell = getActiveCostCellMeta();
    updateCostValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    recomputeCostForecastValues();
    renderCostModule();
    focusCostCellInput(activeCell.rowIndex, activeCell.colIndex);
    return;
  }
  const input = event.target.closest("input[data-cost-code][data-cost-month]");
  if (!input) return;
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    updateCostValue(input.dataset.costCode, input.dataset.costMonth, input.value, { recompute: false });
    recomputeCostForecastValues();
    fillCostSelectionToLastRow();
    render();
    focusCostCellInput(Number(input.dataset.costRow), Number(input.dataset.costCol), { selectText: false });
    return;
  }
  let nextRowIndex = Number(input.dataset.costRow);
  let nextColIndex = Number(input.dataset.costCol);
  if (event.key === "Enter") nextRowIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "Tab") nextColIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "ArrowUp") nextRowIndex -= 1;
  else if (event.key === "ArrowDown") nextRowIndex += 1;
  else if (event.key === "ArrowLeft") {
    if ((input.selectionStart ?? 0) !== 0 || (input.selectionEnd ?? 0) !== 0) return;
    nextColIndex -= 1;
  } else if (event.key === "ArrowRight") {
    const caretEnd = input.value.length;
    if ((input.selectionStart ?? caretEnd) !== caretEnd || (input.selectionEnd ?? caretEnd) !== caretEnd) return;
    nextColIndex += 1;
  } else {
    return;
  }
  event.preventDefault();
  updateCostValue(input.dataset.costCode, input.dataset.costMonth, input.value, { recompute: false });
  recomputeCostForecastValues();
  setActiveCostCell(nextRowIndex, nextColIndex);
}

function updateMarketingValue(code, month, value, options) {
  setMarketingCellRawValue(code, month, value, options);
}

function handleMarketingInput(event) {
  const metaInput = event.target.closest("[data-marketing-meta-field][data-marketing-meta-code]");
  if (metaInput) {
    const row = getMarketingRowByCode(metaInput.dataset.marketingMetaCode);
    if (!row) return;
    row[metaInput.dataset.marketingMetaField] = metaInput.value;
    if (metaInput.dataset.marketingMetaField === "estimateMethodRaw") {
      row.forecastLogicRaw = deriveMarketingForecastLogic(row.estimateMethodRaw, row.forecastLogicRaw);
      if (row.computed) {
        getMarketingMonths().forEach((month) => {
          row.rawValues[month] = row.estimateMethodRaw;
        });
      } else if (isRateEstimateMethod(row.estimateMethodRaw)) {
        getMarketingMonths().forEach((month) => {
          row.rawValues[month] = "";
        });
      }
    }
    if (["referencePeriodRaw", "costDriver", "estimateMethodRaw"].includes(metaInput.dataset.marketingMetaField)) {
      row.suggestedValueText = deriveMarketingSuggestedValueText(row.code, row.referencePeriodRaw);
      if (row.estimateMethodRaw === "费用率" && metaInput.dataset.marketingMetaField !== "forecastParamRaw" && row.suggestedValueText !== "-") {
        row.forecastParamRaw = row.forecastParamRaw || row.suggestedValueText;
      }
    }
    recomputeMarketingForecastValues();
    if (event.type === "change" || metaInput.dataset.marketingMetaField !== "forecastParamRaw") {
      renderMarketingModule();
      if (event.type === "input" && metaInput.dataset.marketingMetaField === "forecastParamRaw") {
        refocusMetaFieldInput(els.marketingModule, `[data-marketing-meta-field="forecastParamRaw"][data-marketing-meta-code="${row.code}"]`);
      }
    } else if (metaInput.dataset.marketingMetaField === "forecastParamRaw") {
      renderMarketingModule();
      refocusMetaFieldInput(els.marketingModule, `[data-marketing-meta-field="forecastParamRaw"][data-marketing-meta-code="${row.code}"]`);
    }
    return;
  }

  const formulaInput = event.target.closest("#marketing-formula-input");
  if (formulaInput) {
    const activeCell = getActiveMarketingCellMeta();
    updateMarketingValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    if (event.type === "change") {
      recomputeMarketingForecastValues();
      render();
    }
    return;
  }

  const input = event.target.closest("input[data-marketing-code][data-marketing-month]");
  if (!input) return;
  updateMarketingValue(input.dataset.marketingCode, input.dataset.marketingMonth, input.value, { recompute: false });
  state.marketingEditCell = { rowIndex: Number(input.dataset.marketingRow), colIndex: Number(input.dataset.marketingCol) };
  const formulaBar = document.getElementById("marketing-formula-input");
  if (formulaBar) formulaBar.value = input.value;
  if (event.type === "change") {
    recomputeMarketingForecastValues();
    state.marketingEditCell = null;
    render();
  }
}

function handleMarketingClick(event) {
  if (state.marketingFormulaInsertLock) {
    state.marketingFormulaInsertLock = false;
    return;
  }

  const fillHandle = event.target.closest("[data-marketing-fill-handle]");
  if (fillHandle) {
    const cell = fillHandle.closest("td[data-marketing-row][data-marketing-col]");
    if (cell && event.detail >= 2) {
      event.preventDefault();
      event.stopPropagation();
      setMarketingSelection(Number(cell.dataset.marketingRow), Number(cell.dataset.marketingCol));
      state.marketingEditCell = { rowIndex: Number(cell.dataset.marketingRow), colIndex: Number(cell.dataset.marketingCol) };
      fillMarketingSelectionToLastRow();
      render();
      focusMarketingCellInput(Number(cell.dataset.marketingRow), Number(cell.dataset.marketingCol), { selectText: false });
      return;
    }
  }

  const action = event.target.closest("[data-marketing-action]");
  if (action?.dataset.marketingAction === "fill-to-end") {
    fillMarketingSelectionToLastRow();
    render();
    focusMarketingCellInput(getActiveMarketingCell().rowIndex, getActiveMarketingCell().colIndex, { selectText: false });
    return;
  }

  const header = event.target.closest("[data-marketing-row-header]");
  if (header) {
    const rowIndex = Number(header.dataset.marketingRowHeader);
    setMarketingSelection(rowIndex, 0, rowIndex, getMarketingMonths().length - 1, "row");
    state.marketingEditCell = null;
    renderMarketingModule();
    return;
  }

  const cell = event.target.closest("td[data-marketing-row][data-marketing-col]");
  if (!cell) return;
  if (tryInsertMarketingFormulaReferenceFromCell(cell)) {
    event.preventDefault();
    return;
  }
  const rowIndex = Number(cell.dataset.marketingRow);
  const colIndex = Number(cell.dataset.marketingCol);
  if (!event.target.closest("input")) {
    setActiveMarketingCell(rowIndex, colIndex, { selectText: true });
  }
}

function handleMarketingPaste(event) {
  const input = event.target.closest("input[data-marketing-code][data-marketing-month]");
  if (!input) return;
  const text = event.clipboardData?.getData("text");
  if (!text) return;

  event.preventDefault();
  const months = getMarketingMonths();
  const startRowIndex = marketingAccountCodes.indexOf(input.dataset.marketingCode);
  const startColIndex = months.indexOf(input.dataset.marketingMonth);
  if (startRowIndex < 0 || startColIndex < 0) return;

  let appliedEndRowIndex = startRowIndex;
  let appliedEndColIndex = startColIndex;
  text.trim().split(/\r?\n/).forEach((line, rowOffset) => {
    const code = marketingAccountCodes[startRowIndex + rowOffset];
    if (!code) return;
    appliedEndRowIndex = startRowIndex + rowOffset;
    line.split("\t").forEach((cell, colOffset) => {
      const month = months[startColIndex + colOffset];
      if (!month) return;
      appliedEndColIndex = Math.max(appliedEndColIndex, startColIndex + colOffset);
      updateMarketingValue(code, month, cell, { recompute: false });
    });
  });

  recomputeMarketingForecastValues();
  setMarketingSelection(startRowIndex, startColIndex, appliedEndRowIndex, appliedEndColIndex);
  state.marketingEditCell = { rowIndex: startRowIndex, colIndex: startColIndex };
  render();
}

function handleMarketingFocusIn(event) {
  const input = event.target.closest("input[data-marketing-code][data-marketing-month]");
  if (!input) return;
  const rowIndex = Number(input.dataset.marketingRow);
  const colIndex = Number(input.dataset.marketingCol);
  state.marketingEditCell = { rowIndex, colIndex };
  setMarketingSelection(rowIndex, colIndex);
  input.value = getMarketingCellEditValue(input.dataset.marketingCode, input.dataset.marketingMonth);
  input.select();
  const formulaBar = document.getElementById("marketing-formula-input");
  if (formulaBar) {
    formulaBar.value = input.value;
    formulaBar.readOnly = isAutoCalculatedForecastRow(getMarketingRowByCode(input.dataset.marketingCode));
  }
}

function handleMarketingCopy(event) {
  if (!event.clipboardData) return;
  if (shouldUseNativeCopy(event)) return;
  event.preventDefault();
  event.clipboardData.setData("text/plain", getMarketingSelectionText());
}

function handleMarketingMouseDown(event) {
  const handle = event.target.closest("[data-marketing-fill-handle]");
  if (handle) {
    event.preventDefault();
    event.stopPropagation();
    const cell = handle.closest("td[data-marketing-row][data-marketing-col]");
    if (cell) {
      const rowIndex = Number(cell.dataset.marketingRow);
      const colIndex = Number(cell.dataset.marketingCol);
      setMarketingSelection(rowIndex, colIndex);
      state.marketingEditCell = { rowIndex, colIndex };
    }
    const selection = normalizeMarketingSelection();
    state.marketingFillDrag = { targetCol: selection.endCol, startCol: selection.endCol, dragging: false };
    return;
  }

  const cell = event.target.closest("td[data-marketing-row][data-marketing-col]");
  if (!cell) return;
  if (event.target.closest("input")) return;
  const rowIndex = Number(cell.dataset.marketingRow);
  const colIndex = Number(cell.dataset.marketingCol);
  setMarketingSelection(rowIndex, colIndex);
  state.marketingEditCell = null;
  renderMarketingModule();
}

function handleMarketingMouseOver(event) {
  if (!state.marketingFillDrag) return;
  const cell = event.target.closest("td[data-marketing-row][data-marketing-col]");
  if (!cell) return;
  const colIndex = Number(cell.dataset.marketingCol);
  if (colIndex > normalizeMarketingSelection().endCol) {
    state.marketingFillDrag.dragging = true;
  }
  state.marketingFillDrag.targetCol = Math.max(colIndex, normalizeMarketingSelection().endCol);
  renderMarketingModule();
}

function handleMarketingDoubleClick(event) {
  const handle = event.target.closest("[data-marketing-fill-handle]");
  if (!handle) return;
  event.preventDefault();
  fillMarketingSelectionToLastRow();
  render();
  focusMarketingCellInput(getActiveMarketingCell().rowIndex, getActiveMarketingCell().colIndex, { selectText: false });
}

function handleMarketingKeyDown(event) {
  const formulaInput = event.target.closest("#marketing-formula-input");
  if (formulaInput && event.key === "Enter") {
    event.preventDefault();
    const activeCell = getActiveMarketingCellMeta();
    updateMarketingValue(activeCell.code, activeCell.month, formulaInput.value, { recompute: false });
    recomputeMarketingForecastValues();
    renderMarketingModule();
    focusMarketingCellInput(activeCell.rowIndex, activeCell.colIndex);
    return;
  }

  const input = event.target.closest("input[data-marketing-code][data-marketing-month]");
  if (!input) return;

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "d") {
    event.preventDefault();
    updateMarketingValue(input.dataset.marketingCode, input.dataset.marketingMonth, input.value, { recompute: false });
    recomputeMarketingForecastValues();
    fillMarketingSelectionToLastRow();
    render();
    focusMarketingCellInput(Number(input.dataset.marketingRow), Number(input.dataset.marketingCol), { selectText: false });
    return;
  }

  let nextRowIndex = Number(input.dataset.marketingRow);
  let nextColIndex = Number(input.dataset.marketingCol);

  if (event.key === "Enter") nextRowIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "Tab") nextColIndex += event.shiftKey ? -1 : 1;
  else if (event.key === "ArrowUp") nextRowIndex -= 1;
  else if (event.key === "ArrowDown") nextRowIndex += 1;
  else if (event.key === "ArrowLeft") {
    if ((input.selectionStart ?? 0) !== 0 || (input.selectionEnd ?? 0) !== 0) return;
    nextColIndex -= 1;
  } else if (event.key === "ArrowRight") {
    const caretEnd = input.value.length;
    if ((input.selectionStart ?? caretEnd) !== caretEnd || (input.selectionEnd ?? caretEnd) !== caretEnd) return;
    nextColIndex += 1;
  } else {
    return;
  }

  event.preventDefault();
  updateMarketingValue(input.dataset.marketingCode, input.dataset.marketingMonth, input.value, { recompute: false });
  recomputeMarketingForecastValues();
  setActiveMarketingCell(nextRowIndex, nextColIndex);
}

async function parsePostFileViaApi(file) {
  const body = await file.arrayBuffer();
  const { response } = await fetchFromApi("/upload-data?scope=post", {
    method: "POST",
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    },
    body
  });
  if (!response.ok) throw new Error("parse_failed");
  return response.json();
}

async function parseFlowFileViaApi(file) {
  const body = await file.arrayBuffer();
  const { response } = await fetchFromApi("/upload-data?scope=flow", {
    method: "POST",
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    },
    body
  });
  if (!response.ok) throw new Error("parse_failed");
  return response.json();
}

async function loadTemplateLayout() {
  try {
    const { response } = await fetchFromApi("/template-layout");
    if (!response.ok) {
      state.templateLayoutRows = normalizeTemplateLayoutRows(defaultTemplateLayoutRows);
      return;
    }
    const payload = await response.json();
    const rows = Array.isArray(payload?.rows) ? payload.rows : [];
    const parsedRows = rows
      .map((row) => ({
        code: String(row.code || "").trim(),
        name: String(row.name || "").trim(),
        enName: String(row.enName || row.en_name || "").trim(),
        indent: Number.isFinite(Number(row.indent)) ? Number(row.indent) : 0
      }))
      .filter((row) => row.code);
    loadedTemplateLayoutNameLookup = new Map(parsedRows.map((row) => [row.code, row.name]).filter((entry) => entry[0] && entry[1]));
    state.templateLayoutRows = normalizeTemplateLayoutRows(parsedRows.length ? parsedRows : defaultTemplateLayoutRows);
  } catch {
    loadedTemplateLayoutNameLookup = new Map();
    state.templateLayoutRows = normalizeTemplateLayoutRows(defaultTemplateLayoutRows);
  }
}

function mapPostRowsToActualRows(rows) {
  return normalizePostRows(rows).map((row) => ({
    month: row.dt,
    code: row.fcst_acc_code,
    account: row.fcst_acc_name,
    updateDt: row.update_dt,
    source: "分摊后",
    amount: row.amount,
    hc: row.hc,
    revenue: row.revenue
  }));
}

function getHistoricalMetricMonthlyMap(codeOrCodes) {
  const targetCodes = new Set(
    (Array.isArray(codeOrCodes) ? codeOrCodes : [codeOrCodes])
      .map((code) => String(code || "").trim())
      .filter(Boolean)
  );
  const grouped = new Map();

  getHistoricalPostRows()
    .filter((row) => targetCodes.has(String(row.code || "").trim()))
    .forEach((row) => {
      const month = normalizeMonth(row.month || row.dt);
      if (!month) return;
      if (!grouped.has(month)) grouped.set(month, []);
      grouped.get(month).push(row);
    });

  const monthlyMap = new Map();
  grouped.forEach((items, month) => {
    const latestUpdate = items.reduce((latest, item) => {
      const value = String(item.updateDt || "");
      return value > latest ? value : latest;
    }, "");
    const selected = latestUpdate ? items.filter((item) => String(item.updateDt || "") === latestUpdate) : items;
    monthlyMap.set(month, sum(selected.map((item) => Number(item.amount || 0))));
  });
  return monthlyMap;
}

function getHistoricalMetricRangeTotal(codeOrCodes, startMonth, endMonth = startMonth, mode = "cumulative") {
  const normalizedStart = normalizeMonth(startMonth);
  const normalizedEnd = normalizeMonth(endMonth || startMonth);
  if (!normalizedStart || !normalizedEnd) return 0;
  const [rangeStart, rangeEnd] = normalizedStart <= normalizedEnd ? [normalizedStart, normalizedEnd] : [normalizedEnd, normalizedStart];
  const monthlyMap = getHistoricalMetricMonthlyMap(codeOrCodes);
  const values = [...monthlyMap.entries()]
    .filter(([month]) => month >= rangeStart && month <= rangeEnd)
    .map(([, value]) => Number(value || 0));
  if (!values.length) return 0;
  if (mode === "average") {
    return avg(values.map((value) => Math.abs(value)));
  }
  return Math.abs(sum(values));
}

function refreshHistoricalPostPivot(fallbackDisplayName = state.postPivot.displayName || "演示项目") {
  const postRows = getHistoricalPostRows();
  state.postPivot = buildPostPivot(postRows, fallbackDisplayName);
}

function loadPivotFromPostRows(rows, fallbackDisplayName) {
  state.actualRows.post = mapPostRowsToActualRows(rows);
  state.actualImportFlags.post = true;
  if (!state.actualRows.pre.length && !state.actualRows.allocated.length) {
    state.importMode = "post";
    state.ruleScope = "post";
  }
  refreshHistoricalPostPivot(fallbackDisplayName);
  applyBoundaryDefaults();
  refreshRules();
  render();
}

function getUploadScopeByKind(kind) {
  if (kind === "allocated") return "allocation";
  return kind;
}

function loadActualRowsByScope(scope, rows) {
  if (state.importMode !== "split") {
    if (!state.actualImportFlags.post) {
      state.actualRows.post = [];
    }
    state.actualRows.pre = [];
    state.actualRows.allocated = [];
  }
  state.actualRows[scope] = mapPostRowsToActualRows(rows).map((row) => ({
    ...row,
    source: scope === "allocated" ? "分摊值" : "项目组直属"
  }));
  state.actualImportFlags[scope] = true;
  state.importMode = "split";
  if (state.ruleScope === "post") {
    state.ruleScope = "pre";
  }
  if (!state.actualRows.post.length || scope !== "post") {
    refreshHistoricalPostPivot(state.postPivot.displayName || "演示项目");
  }
  applyBoundaryDefaults();
  refreshRules();
  render();
}

async function handleFileSelect(kind, file) {
  if (!file) return;
  const metaMap = {
    post: els.pnlPostMeta,
    pre: els.pnlPreMeta,
    allocated: els.pnlAllocMeta,
    flow: els.flowMeta
  };
  const targetMeta = metaMap[kind];
  if (!targetMeta) return;

  if (kind === "post") {
    targetMeta.textContent = `已选择 ${file.name}，正在解析...`;
    try {
      const parsed = await parsePostFileViaApi(file);
      const rows = Array.isArray(parsed?.rows) ? parsed.rows : [];
      if (!rows.length) throw new Error("empty_rows");
      const displayName = parsed.displayName || parsed.display_name || file.name.replace(/\.[^.]+$/, "");
      loadPivotFromPostRows(rows, displayName);
      state.apiConnected = true;
      targetMeta.textContent = `已导入 ${file.name}，共 ${rows.length} 条分摊后记录，来源 ${parsed.sheetCount || parsed.sheets?.length || 1} 个 sheet，已刷新历史概览与透视表。`;
      return;
    } catch {
      targetMeta.textContent = `已选择 ${file.name}，但当前无法解析该 Excel，页面继续使用现有分摊后数据。请确认本地解析服务已启动。`;
      render();
      return;
    }
  }

  if (kind === "flow") {
    targetMeta.textContent = `已选择 ${file.name}，正在解析流水...`;
    try {
      const parsed = await parseFlowFileViaApi(file);
      const rows = Array.isArray(parsed?.rows) ? parsed.rows : [];
      if (!rows.length) throw new Error("empty_rows");
      applyImportedFlowRows(rows, file.name, parsed);
      state.apiConnected = true;
      targetMeta.textContent = `已导入 ${file.name}，共 ${state.flowImportRowCount} 个月；步骤 2.2 的 A2 已按月回填，A3 / A4 / A5 已同步联动刷新。`;
      render();
      return;
    } catch {
      targetMeta.textContent = `已选择 ${file.name}，但当前无法解析该流水 Excel，页面继续使用${state.flowImportName ? "上一次导入的流水" : "演示流水"}。请确认本地解析服务已启动。`;
      render();
      return;
    }
  }

  if (kind === "pre" || kind === "allocated") {
    const uploadScope = getUploadScopeByKind(kind);
    const stateScope = kind;
    const label = kind === "pre" ? "分摊前" : "分摊值";
    targetMeta.textContent = `已选择 ${file.name}，正在解析${label}数据...`;
    try {
      const body = await file.arrayBuffer();
      const { response } = await fetchFromApi(`/upload-data?scope=${uploadScope}`, {
        method: "POST",
        headers: {
          "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        },
        body
      });
      if (!response.ok) throw new Error("parse_failed");
      const parsed = await response.json();
      const rows = Array.isArray(parsed?.rows) ? parsed.rows : [];
      if (!rows.length) throw new Error("empty_rows");
      loadActualRowsByScope(stateScope, rows);
      state.apiConnected = true;
      targetMeta.textContent = `已导入 ${file.name}，共 ${rows.length} 条${label}记录，来源 ${parsed.sheetCount || parsed.sheets?.length || 1} 个 sheet，已联动刷新历史概览和后续步骤口径。`;
      return;
    } catch {
      targetMeta.textContent = `已选择 ${file.name}，但当前无法解析${label} Excel，页面继续使用现有${label}数据。请确认本地解析服务已启动。`;
      render();
      return;
    }
  }

  targetMeta.textContent = `已选择 ${file.name}，当前版本等待后端解析接入。`;
}

async function checkApiHealth() {
  try {
    const { response, base } = await fetchFromApi("/health");
    state.apiConnected = response.ok;
    state.apiBase = base;
    state.apiStatusText = response.ok ? "已连接，可读取真实 Excel" : (runtimeContext.isFileMode ? "文件模式，解析服务异常" : "服务响应异常，当前使用演示数据");
    state.apiStatusHint = response.ok
      ? `已连接到 ${state.apiBase}，可直接解析上传的 Excel 文件。`
      : `已访问 ${state.apiBase}，但健康检查未通过，当前回退为演示数据。`;
  } catch (error) {
    state.apiConnected = false;
    const message = String(error?.message || error || "");
    const isRefused = /Failed to fetch|NetworkError|Load failed|fetch/i.test(message);
    state.apiStatusText = runtimeContext.isFileMode ? "文件模式" : "未连接，当前使用演示数据";
    state.apiStatusHint = isRefused
      ? (runtimeContext.isFileMode
        ? `当前从本地文件打开；已尝试本机解析服务地址但仍未连通。请确认 4174 解析服务已启动。`
        : `未能访问 ${state.apiBase}。请先启动本地解析服务，然后保持服务窗口运行。`)
      : `本地解析服务检查失败：${message || "未知错误"}。当前继续使用演示数据。`;
  }
  els.apiDot.className = `status-dot ${state.apiConnected ? "online" : runtimeContext.isFileMode ? "muted" : "offline"}`;
  els.apiText.textContent = state.apiStatusText;
  if (els.serverHint) {
    els.serverHint.textContent = state.apiStatusHint;
    els.serverHint.classList.toggle("helper-note-warning", !state.apiConnected && !runtimeContext.isFileMode);
    els.serverHint.classList.toggle("helper-note-success", state.apiConnected || runtimeContext.isFileMode);
  }
}

function startApiHealthPolling() {
  if (runtimeContext.isFileMode) return;
  window.setInterval(async () => {
    const wasConnected = state.apiConnected;
    await checkApiHealth();
    if (!wasConnected && state.apiConnected) {
      await loadTemplateLayout();
      render();
    }
  }, 5000);
}

async function bootstrap() {
  state.persistence.available = Boolean(safeLocalStorage());
  const existingDraftCollection = readDraftCollection();
  if (!state.persistence.available) {
    setDraftStatus("当前浏览器无法使用本地草稿保存", { hasDraft: false, lastSavedAt: "" });
  } else if (existingDraftCollection?.drafts?.length) {
    state.persistence.drafts = existingDraftCollection.drafts;
    state.persistence.activeDraftId = existingDraftCollection.activeDraftId || existingDraftCollection.drafts[0]?.id || "";
    state.persistence.activeDraftName = getDraftEntryById(state.persistence.activeDraftId)?.name || getDefaultDraftName();
    setDraftStatus(`检测到 ${existingDraftCollection.drafts.length} 个本地草稿槽位`, {
      hasDraft: true,
      lastSavedAt: getDraftEntryById(state.persistence.activeDraftId)?.savedAt || existingDraftCollection.drafts[0]?.savedAt || ""
    });
  } else {
    state.persistence.activeDraftName = getDefaultDraftName();
    setDraftStatus("当前还没有本地草稿，编辑后会自动保存到当前槽位", { hasDraft: false, lastSavedAt: "" });
  }
  await checkApiHealth();
  if (state.apiConnected || runtimeContext.isFileMode) await loadTemplateLayout();
  applyBoundaryDefaults();
  refreshRules();
  refreshHistoricalPostPivot("Aniimo");
  if (!state.persistence.activeDraftName) {
    state.persistence.activeDraftName = getDefaultDraftName();
  }
  if (existingDraftCollection?.drafts?.length) {
    const activeDraft = getDraftEntryById(state.persistence.activeDraftId) || existingDraftCollection.drafts[0];
    if (activeDraft?.payload) {
      restoreDraftPayload(activeDraft.payload);
    }
    setDraftStatus(`已加载草稿「${activeDraft?.name || "未命名草稿"}」：${formatTimestamp(activeDraft?.savedAt || "")}`, {
      hasDraft: true,
      lastSavedAt: activeDraft?.savedAt || ""
    });
  } else if (state.persistence.available) {
    state.persistence.lastSavedSignature = getDraftSignature();
    state.persistence.dirty = false;
  }
  syncInputs();
  bindEvents();
  render();
  startApiHealthPolling();
}

bootstrap();


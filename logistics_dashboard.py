import React, { useState, useMemo } from "react";
import {
  ResponsiveContainer, ComposedChart, LineChart, Line, Area, Bar, BarChart,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ReferenceArea, Dot,
} from "recharts";

/* ---------------------------------------------------------------------
   Data (aggregated from data.xlsx: 08_물류비 항목 / 04_출고 / 03_입고 / 05_불량관리)
--------------------------------------------------------------------- */
const COST = [
  { ym: "2024-01", overseas: 921825, online: 13634940, contractor: 0 },
  { ym: "2024-02", overseas: 1660140, online: 14666030, contractor: 0 },
  { ym: "2024-03", overseas: 1662661, online: 19891010, contractor: 0 },
  { ym: "2024-04", overseas: 4035529, online: 27590520, contractor: 0 },
  { ym: "2024-05", overseas: 7957953, online: 24332930, contractor: 0 },
  { ym: "2024-06", overseas: 10171095, online: 34381440, contractor: 2435360 },
  { ym: "2024-07", overseas: 10010104, online: 42477958, contractor: 18817097 },
  { ym: "2024-08", overseas: 9149989, online: 36521320, contractor: 9909229 },
  { ym: "2024-09", overseas: 9264564, online: 28359700, contractor: 8494882 },
  { ym: "2024-10", overseas: 9218135, online: 29070152, contractor: 12381358 },
  { ym: "2024-11", overseas: 9612365, online: 31767610, contractor: 20540839 },
  { ym: "2024-12", overseas: 17348944, online: 34260570, contractor: 19207337 },
  { ym: "2025-01", overseas: 17216664, online: 41108610, contractor: 24330416 },
  { ym: "2025-02", overseas: 17755568, online: 44819640, contractor: 33934979 },
  { ym: "2025-03", overseas: 15966300, online: 40679210, contractor: 43910571 },
  { ym: "2025-04", overseas: 14937324, online: 35281090, contractor: 63704928 },
  { ym: "2025-05", overseas: 13636656, online: 32485220, contractor: 73057822 },
  { ym: "2025-06", overseas: 16578729, online: 35366240, contractor: 94382391 },
  { ym: "2025-07", overseas: 17998222, online: 39295870, contractor: 89222668 },
  { ym: "2025-08", overseas: 18295239, online: 37004230, contractor: 92104779 },
  { ym: "2025-09", overseas: 17266903, online: 30485600, contractor: 82095382 },
  { ym: "2025-10", overseas: 8714007, online: 25818672, contractor: 71811463 },
  { ym: "2025-11", overseas: 13832775, online: 26100940, contractor: 67640447 },
  { ym: "2025-12", overseas: 16337050, online: 31462320, contractor: 63619621 },
  { ym: "2026-01", overseas: 17201942, online: 41687190, contractor: 77457479 },
  { ym: "2026-02", overseas: 11953120, online: 39680705, contractor: 62687955 },
  { ym: "2026-03", overseas: 18281526, online: 41787800, contractor: 77876783 },
  { ym: "2026-04", overseas: 21120706, online: 43281465, contractor: 65547194 },
  { ym: "2026-05", overseas: 13841450, online: 38507985, contractor: 72738693 },
  { ym: "2026-06", overseas: 16590423, online: 33480243, contractor: 69086717 },
  { ym: "2026-07", overseas: 80760713, online: 32347030, contractor: 7228900, flagged: true },
  { ym: "2026-08", overseas: 82805842, online: 30194490, contractor: 7956810, flagged: true },
];

const VOLUME = [
  { ym: "2026-01", overseas: 10268, online: 18302, contractor: 472067 },
  { ym: "2026-02", overseas: 11089, online: 17153, contractor: 347070 },
  { ym: "2026-03", overseas: 6219, online: 18766, contractor: 389230 },
  { ym: "2026-04", overseas: 9729, online: 18947, contractor: 500286 },
  { ym: "2026-05", overseas: 9842, online: 16480, contractor: 280481 },
  { ym: "2026-06", overseas: 25659, online: 13601, contractor: 329302 },
  { ym: "2026-07", overseas: 27466, online: 13038, contractor: 298931 },
  { ym: "2026-08", overseas: 24697, online: 15316, contractor: 457671 },
];

const UNIT = COST.filter((c) => VOLUME.some((v) => v.ym === c.ym)).map((c) => {
  const v = VOLUME.find((v) => v.ym === c.ym);
  return {
    ym: c.ym,
    overseas: +(c.overseas / v.overseas).toFixed(0),
    online: +(c.online / v.online).toFixed(0),
    contractor: +(c.contractor / v.contractor).toFixed(0),
    flagged: c.flagged,
  };
});

const INK = "#22262F";
const INK_SOFT = "#6B7280";
const PAPER = "#F5F4F0";
const CARD = "#FFFFFF";
const LINE = "#E4E1D8";
const OVERSEAS = "#2E6F63";
const ONLINE = "#35507E";
const CONTRACTOR = "#AD5D2E";
const WARN = "#B33A3A";
const WARN_BG = "#FBEDEA";

const FONT = "'Pretendard Variable', Pretendard, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif";

const won = (n) => "₩" + Math.round(n).toLocaleString("ko-KR");
const wonShort = (n) => {
  const v = Math.round(n);
  if (Math.abs(v) >= 100000000) return (v / 100000000).toFixed(1) + "억";
  if (Math.abs(v) >= 10000) return (v / 10000).toFixed(0) + "만";
  return v.toLocaleString("ko-KR");
};

const CHANNELS = [
  { key: "overseas", label: "해외", sub: "DHL + EMS", color: OVERSEAS },
  { key: "online", label: "온라인", sub: "CJ + 카카오", color: ONLINE },
  { key: "contractor", label: "도급사", sub: "오프라인+일본+협찬+샘플+입고+불량", color: CONTRACTOR },
];

function Card({ children, style }) {
  return (
    <div
      style={{
        background: CARD,
        border: `1px solid ${LINE}`,
        borderRadius: 10,
        padding: "20px 22px",
        ...style,
      }}
    >
      {children}
    </div>
  );
}

function KpiCard({ ch }) {
  const last = COST[COST.length - 1];
  const prev = COST[COST.length - 2];
  const cur = last[ch.key];
  const prevVal = prev[ch.key];
  const mom = prevVal ? ((cur - prevVal) / prevVal) * 100 : 0;
  const up = mom >= 0;
  return (
    <Card style={{ position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", left: 0, top: 0, bottom: 0, width: 4, background: ch.color }} />
      <div style={{ paddingLeft: 8 }}>
        <div style={{ fontSize: 12.5, color: INK_SOFT, fontWeight: 600, letterSpacing: 0.2 }}>
          {ch.label} <span style={{ fontWeight: 400 }}>· {ch.sub}</span>
        </div>
        <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginTop: 8 }}>
          <div style={{ fontSize: 26, fontWeight: 700, color: INK, fontVariantNumeric: "tabular-nums" }}>
            {won(cur)}
          </div>
          {last.flagged && (
            <span
              title="원본 데이터 이상치 의심 구간"
              style={{ fontSize: 12, color: WARN, fontWeight: 700, cursor: "help" }}
            >
              ⚠
            </span>
          )}
        </div>
        <div style={{ fontSize: 12.5, color: up ? "#B33A3A" : "#2E6F63", marginTop: 6, fontWeight: 600 }}>
          전월대비 {up ? "▲" : "▼"} {Math.abs(mom).toFixed(1)}% · 2026년 8월 기준
        </div>
      </div>
    </Card>
  );
}

function CustomTooltip({ active, payload, label, unit }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div
      style={{
        background: "#fff",
        border: `1px solid ${LINE}`,
        borderRadius: 8,
        padding: "10px 14px",
        fontSize: 12.5,
        boxShadow: "0 4px 16px rgba(0,0,0,0.08)",
        fontFamily: FONT,
      }}
    >
      <div style={{ fontWeight: 700, marginBottom: 6, color: INK }}>{label}</div>
      {payload.map((p) => (
        <div key={p.dataKey} style={{ display: "flex", justifyContent: "space-between", gap: 16, color: INK_SOFT }}>
          <span style={{ color: p.color, fontWeight: 600 }}>{p.name}</span>
          <span style={{ fontVariantNumeric: "tabular-nums", color: INK }}>
            {unit === "unit" ? won(p.value) + "/개" : won(p.value)}
          </span>
        </div>
      ))}
    </div>
  );
}

export default function LogisticsCostDashboard() {
  const [range, setRange] = useState("all"); // all | y2025_26 | y2026
  const chartData = useMemo(() => {
    if (range === "y2026") return COST.filter((c) => c.ym.startsWith("2026"));
    if (range === "y2025_26") return COST.filter((c) => c.ym >= "2025-01");
    return COST;
  }, [range]);

  const totals = useMemo(() => {
    const t = { overseas: 0, online: 0, contractor: 0 };
    COST.forEach((c) => {
      t.overseas += c.overseas;
      t.online += c.online;
      t.contractor += c.contractor;
    });
    return t;
  }, []);
  const grandTotal = totals.overseas + totals.online + totals.contractor;

  return (
    <div style={{ fontFamily: FONT, background: PAPER, minHeight: "100%", padding: "28px 28px 40px" }}>
      {/* Header */}
      <div style={{ marginBottom: 22 }}>
        <div style={{ fontSize: 22, fontWeight: 700, color: INK, letterSpacing: -0.3 }}>물류비 운영 대시보드</div>
        <div style={{ fontSize: 13, color: INK_SOFT, marginTop: 6, lineHeight: 1.6 }}>
          비용 2024-01 ~ 2026-08 · 물동량·단가 2026-01 ~ 2026-08 (2024~2025년 물동량 데이터 없음) ·{" "}
          <span style={{ color: INK }}>퀵/화물, 고고밴(밀크런)은 범위 제외</span>
        </div>
      </div>

      {/* KPI row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14, marginBottom: 16 }}>
        {CHANNELS.map((ch) => (
          <KpiCard key={ch.key} ch={ch} />
        ))}
      </div>

      {/* Anomaly banner */}
      <div
        style={{
          background: WARN_BG,
          border: `1px solid ${WARN}33`,
          borderRadius: 10,
          padding: "12px 16px",
          fontSize: 12.5,
          color: WARN,
          marginBottom: 20,
          lineHeight: 1.6,
        }}
      >
        <b>⚠ 데이터 확인 필요</b> — 2026-07, 2026-08 두 달은 해외(DHL/EMS)와 도급사 값이 서로 뒤바뀐 것으로 보입니다.
        도급사 단가가 갑자기 20원/개대로 떨어지고 해외 비용이 8천만원대로 급등하는 패턴이 그 근거입니다. 아래 차트에서
        점선·강조 표시된 구간이며, 원본 08_물류비 항목 시트를 재확인해 주세요.
      </div>

      {/* Main trend chart */}
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
          <div style={{ fontSize: 14.5, fontWeight: 700, color: INK }}>월별 물류비 추이</div>
          <div style={{ display: "flex", gap: 6 }}>
            {[
              { key: "all", label: "전체 (2024~)" },
              { key: "y2025_26", label: "2025~2026" },
              { key: "y2026", label: "2026년" },
            ].map((r) => (
              <button
                key={r.key}
                onClick={() => setRange(r.key)}
                style={{
                  fontFamily: FONT,
                  fontSize: 12,
                  padding: "5px 12px",
                  borderRadius: 999,
                  border: `1px solid ${range === r.key ? INK : LINE}`,
                  background: range === r.key ? INK : "#fff",
                  color: range === r.key ? "#fff" : INK_SOFT,
                  cursor: "pointer",
                  fontWeight: 600,
                }}
              >
                {r.label}
              </button>
            ))}
          </div>
        </div>
        <ResponsiveContainer width="100%" height={300}>
          <ComposedChart data={chartData} margin={{ top: 16, right: 12, left: 4, bottom: 0 }}>
            <CartesianGrid stroke={LINE} vertical={false} />
            <XAxis dataKey="ym" tick={{ fontSize: 11, fill: INK_SOFT }} axisLine={{ stroke: LINE }} tickLine={false} />
            <YAxis
              tickFormatter={wonShort}
              tick={{ fontSize: 11, fill: INK_SOFT }}
              axisLine={false}
              tickLine={false}
              width={54}
            />
            <Tooltip content={<CustomTooltip />} />
            <Legend wrapperStyle={{ fontSize: 12 }} />
            {chartData.some((d) => d.flagged) && (
              <ReferenceArea
                x1={chartData.find((d) => d.flagged)?.ym}
                x2={chartData[chartData.length - 1]?.ym}
                fill={WARN}
                fillOpacity={0.06}
              />
            )}
            <Line type="monotone" dataKey="overseas" name="해외" stroke={OVERSEAS} strokeWidth={2.4} dot={false} />
            <Line type="monotone" dataKey="online" name="온라인" stroke={ONLINE} strokeWidth={2.4} dot={false} />
            <Line type="monotone" dataKey="contractor" name="도급사" stroke={CONTRACTOR} strokeWidth={2.4} dot={false} />
          </ComposedChart>
        </ResponsiveContainer>
      </Card>

      {/* Two column: share + unit cost */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1.4fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <div style={{ fontSize: 14.5, fontWeight: 700, color: INK, marginBottom: 12 }}>
            누적 비용 비중 <span style={{ fontWeight: 400, color: INK_SOFT, fontSize: 12 }}>(2024-01~2026-08)</span>
          </div>
          {CHANNELS.map((ch) => {
            const pct = (totals[ch.key] / grandTotal) * 100;
            return (
              <div key={ch.key} style={{ marginBottom: 14 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 5 }}>
                  <span style={{ color: INK, fontWeight: 600 }}>{ch.label}</span>
                  <span style={{ color: INK_SOFT, fontVariantNumeric: "tabular-nums" }}>
                    {won(totals[ch.key])} · {pct.toFixed(1)}%
                  </span>
                </div>
                <div style={{ height: 8, borderRadius: 6, background: "#EEECE5", overflow: "hidden" }}>
                  <div style={{ width: `${pct}%`, height: "100%", background: ch.color, borderRadius: 6 }} />
                </div>
              </div>
            );
          })}
          <div style={{ fontSize: 11.5, color: INK_SOFT, marginTop: 10, borderTop: `1px solid ${LINE}`, paddingTop: 10 }}>
            총 누적 물류비 <b style={{ color: INK }}>{won(grandTotal)}</b>
          </div>
        </Card>

        <Card>
          <div style={{ fontSize: 14.5, fontWeight: 700, color: INK, marginBottom: 4 }}>
            2026년 채널별 단가 추이 <span style={{ fontWeight: 400, color: INK_SOFT, fontSize: 12 }}>(비용 ÷ 물동량, 원/개)</span>
          </div>
          <ResponsiveContainer width="100%" height={252}>
            <BarChart data={UNIT} margin={{ top: 16, right: 12, left: 4, bottom: 0 }} barGap={4}>
              <CartesianGrid stroke={LINE} vertical={false} />
              <XAxis dataKey="ym" tick={{ fontSize: 11, fill: INK_SOFT }} axisLine={{ stroke: LINE }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: INK_SOFT }} axisLine={false} tickLine={false} width={40} />
              <Tooltip content={<CustomTooltip unit="unit" />} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="overseas" name="해외" fill={OVERSEAS} radius={[3, 3, 0, 0]} />
              <Bar dataKey="online" name="온라인" fill={ONLINE} radius={[3, 3, 0, 0]} />
              <Bar dataKey="contractor" name="도급사" fill={CONTRACTOR} radius={[3, 3, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <div style={{ fontSize: 11, color: WARN, marginTop: 4 }}>
            ⚠ 2026-07, 08월은 데이터 이상치 의심 구간으로 단가 신뢰도가 낮습니다.
          </div>
        </Card>
      </div>

      <div style={{ fontSize: 11, color: INK_SOFT, textAlign: "right" }}>
        기준: 해외=DHL+EMS · 온라인=CJ+카카오 · 도급사=오프라인+일본+협찬+샘플 출고, 입고, 불량 합산 · 퀵/화물·고고밴 제외
      </div>
    </div>
  );
}

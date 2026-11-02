/* Hiển thị chữ + công thức (KaTeX) và dựng bản in. Dùng chung cho index.html và print.html. */
const TOKEN = /(\$[^$]+\$|\*\*.+?\*\*|\*[^*\s][^*]*?\*)/;
const esc = s => s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
function math(latex) {
  try { if (window.katex) return katex.renderToString(latex, { throwOnError:false, output:"html" }); } catch(e){}
  return "<code>" + esc(latex) + "</code>";
}
function rich(text) {
  return (text || "").split("\n").map(line => {
    let boxed = false;
    if (line.startsWith("[box]")) { boxed = true; line = line.slice(5).trim(); }
    const html = line.split(TOKEN).map(p => {
      if (!p) return "";
      if (p.length > 2 && p[0] === "$" && p.endsWith("$")) return math(p.slice(1,-1));
      if (p.length > 4 && p.startsWith("**") && p.endsWith("**")) return "<b>" + esc(p.slice(2,-2)) + "</b>";
      if (p.length > 2 && p[0] === "*" && p.endsWith("*")) return "<i>" + esc(p.slice(1,-1)) + "</i>";
      return esc(p);
    }).join("");
    return boxed ? '<span class="box">' + html + "</span>" : html;
  }).join("<br>");
}

/* Một khối -> HTML để in. job = mã công việc, n = số trang gốc (để cắt hình). */
function printBlock(job, n, b) {
  const al = a => ({ left:"left", center:"center", right:"right", justify:"justify" }[a] || "left");
  const img = (bb, style) => '<img src="/api/job/' + job + "/crop/" + n + ".jpg?bbox=" + bb.join(",") + '" style="' + style + '">';
  const figW = bb => Math.min(17, Math.max(3, (bb[2] - bb[0]) / 100 * 21));
  switch (b.type) {
    case "paragraph":
      return '<p style="text-align:' + al(b.align) + ";margin-left:" + (b.indent || 0) + "cm;text-indent:" + (b.first_line ? "1cm" : "0") + '">' + rich(b.text) + "</p>";
    case "boxed": return '<p style="text-align:center"><span class="pbox">' + rich(b.text) + "</span></p>";
    case "formula": return '<p style="text-align:center">' + math(b.latex || "") + "</p>";
    case "figure": return '<p style="text-align:center">' + img(b.bbox, "width:" + figW(b.bbox) + "cm") + "</p>";
    case "columns": {
      const w = (b.widths && b.widths.length === b.cells.length) ? b.widths : b.cells.map(() => 100 / b.cells.length);
      const tot = w.reduce((x, y) => x + y, 0) || 1;
      return '<table class="pcols"><tr>' + b.cells.map((c, i) => '<td style="width:' + (100 * w[i] / tot) + "%;text-align:" + al((b.aligns || [])[i]) + '">' + rich(c) + "</td>").join("") + "</tr></table>";
    }
    case "text_figure": {
      const fw = Math.min(9.5, Math.max(4, figW(b.bbox)));
      const t = '<td style="text-align:' + al(b.align) + '">' + rich(b.text) + "</td>";
      const f = '<td style="width:' + fw + 'cm;text-align:center">' + img(b.bbox, "width:100%") + "</td>";
      return '<table class="pcols"><tr>' + (b.side === "left" ? f + t : t + f) + "</tr></table>";
    }
  }
  return "";
}

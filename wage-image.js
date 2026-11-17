// 生成工资单 / 汇总 / 工时图片（1080×1350）。
const moneyFmt = new Intl.NumberFormat('zh-CN', {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});
const yuan = (value) => `¥${moneyFmt.format(Math.round((Number(value) || 0) * 100) / 100)}`;

const FONT = '"PingFang SC", -apple-system, system-ui, sans-serif';
const GREEN = '#2e6b4f';
const INK = '#191a18';
const MUTED = '#79776f';
const PAY = '#2f6f8f';
const OT = '#9a5b00';

function roundRect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.roundRect(x, y, w, h, r);
}

export function buildStatementImage({ title, subtitle, fields = {}, sections, totals, generatedAt }) {
  const single = sections.length === 1;
  const only = sections[0] || {};
  const show = {
    days: fields.days !== false,
    overtime: fields.overtime !== false,
    wage: fields.wage !== false,
    paid: fields.paid !== false,
    balance: fields.balance !== false,
  };

  const canvas = document.createElement('canvas');
  canvas.width = 1080;
  canvas.height = 1350;
  const ctx = canvas.getContext('2d');
  if (!ctx) return Promise.reject(new Error('no canvas'));

  ctx.fillStyle = '#f3f2ee';
  ctx.fillRect(0, 0, 1080, 1350);

  const gradient = ctx.createLinearGradient(0, 0, 1080, 500);
  gradient.addColorStop(0, '#31785a');
  gradient.addColorStop(1, '#1c4936');
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, 1080, 500);

  ctx.fillStyle = 'rgba(255,255,255,0.72)';
  ctx.font = `600 30px ${FONT}`;
  ctx.textAlign = 'left';
  const brand = show.wage || show.paid || show.balance
    ? single
      ? '记工 · 工资单'
      : '记工 · 汇总'
    : single
      ? '记工 · 工时'
      : '记工 · 工时汇总';
  ctx.fillText(brand, 72, 116);

  const name = title || '未命名项目';
  let nameSize = 56;
  ctx.font = `700 ${nameSize}px ${FONT}`;
  while (ctx.measureText(name).width > 900 && nameSize > 30) {
    nameSize -= 2;
    ctx.font = `700 ${nameSize}px ${FONT}`;
  }
  ctx.fillStyle = '#ffffff';
  ctx.fillText(name, 72, 236);

  ctx.fillStyle = 'rgba(255,255,255,0.85)';
  ctx.font = `500 36px ${FONT}`;
  ctx.fillText(subtitle || '', 72, 324);

  const label = (text, baseline, size = 32, color = MUTED) => {
    ctx.fillStyle = color;
    ctx.font = `${size}px ${FONT}`;
    ctx.textAlign = 'left';
    ctx.fillText(text, 132, baseline);
  };
  const value = (text, baseline, size = 38, color = INK, weight = 600) => {
    ctx.fillStyle = color;
    ctx.font = `${weight} ${size}px ${FONT}`;
    ctx.textAlign = 'right';
    ctx.fillText(text, 948, baseline);
  };
  const divider = (y) => {
    ctx.strokeStyle = 'rgba(35, 36, 33, 0.10)';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(132, y);
    ctx.lineTo(948, y);
    ctx.stroke();
  };

  // 主数字：有余额就显示余额，没有就显示工时/加班，再退到工资。
  let heroLabel;
  let heroValue;
  if (show.balance) {
    heroLabel = single ? '本期应得余额' : '合计应得余额';
    heroValue = yuan(single ? only.wage - only.paid : totals.wage - totals.paid);
  } else if (show.wage) {
    heroLabel = single ? '本期应得工资' : '合计应得工资';
    heroValue = yuan(single ? only.wage : totals.wage);
  } else if (show.paid) {
    heroLabel = single ? '本期支账' : '合计支账';
    heroValue = yuan(single ? only.paid : totals.paid);
  } else if (show.days) {
    heroLabel = single ? '本期工日' : '合计工日';
    heroValue = `${single ? only.days : totals.days} 天`;
  } else if (show.overtime) {
    heroLabel = single ? '本期加班' : '合计加班';
    heroValue = `${single ? only.hours : totals.hours} 小时`;
  } else {
    heroLabel = '本期';
    heroValue = '';
  }

  const workMode = !show.wage && !show.paid && !show.balance;
  const rowsShown = single ? 0 : sections.slice(0, 4);
  const singleRows = [];
  if (single) {
    if (workMode) {
      if (show.days && show.overtime) singleRows.push({ label: '加班', value: `${only.hours} 小时` });
      singleRows.push({ label: '上班', value: `${only.worked || 0} 天` });
      singleRows.push({ label: '休息', value: `${only.rested || 0} 天` });
    } else {
      if (show.days) singleRows.push({ label: '工日', value: `${only.days} 天` });
      if (show.overtime) singleRows.push({ label: '加班', value: `${only.hours} 小时` });
      if (show.wage) singleRows.push({ label: '应得工资', value: yuan(only.wage) });
      if (show.paid) singleRows.push({ label: '支账', value: yuan(only.paid) });
    }
  }
  let cardHeight;
  if (single) {
    cardHeight = 250 + singleRows.length * 78 + (show.balance ? 200 : 40);
  } else {
    cardHeight =
      502 +
      56 * rowsShown.length +
      (sections.length > rowsShown.length ? 44 : 0) +
      (show.wage || show.paid ? 0 : 0);
  }

  const cardTop = Math.max(420, 420 + Math.round((740 - cardHeight) / 2));

  ctx.save();
  ctx.shadowColor = 'rgba(20, 60, 44, 0.16)';
  ctx.shadowBlur = 40;
  ctx.shadowOffsetY = 18;
  ctx.fillStyle = '#ffffff';
  roundRect(ctx, 72, cardTop, 936, cardHeight, 44);
  ctx.fill();
  ctx.restore();

  label(heroLabel, cardTop + 86, 30);
  ctx.font = `700 88px ${FONT}`;
  const heroSize = ctx.measureText(heroValue).width > 780 ? 64 : 88;
  value(heroValue, cardTop + 200, heroSize, GREEN, 700);

  divider(cardTop + 250);

  if (single) {
    const rows = singleRows;
    rows.forEach((item, index) => {
      const baseline = cardTop + 250 + 78 * (index + 1);
      label(item.label, baseline);
      value(item.value, baseline);
    });
    if (show.balance) {
      const afterRows = cardTop + 250 + 78 * rows.length + 34;
      divider(afterRows);
      label('项目累计余额', afterRows + 76, 32, INK);
      value(yuan(only.allWage - only.allPaid), afterRows + 76, 42, INK, 700);
    }
  } else {
    const cols = [];
    if (show.wage || show.paid) {
      if (show.wage) cols.push({ key: 'wage', label: '应得' });
      if (show.paid) cols.push({ key: 'paid', label: '支账' });
    } else {
      if (show.days) cols.push({ key: 'days', label: '工日' });
      if (show.overtime) cols.push({ key: 'hours', label: '加班' });
    }
    const xs = cols.length === 1 ? [948] : cols.map((_, index) => 948 - (cols.length - 1 - index) * 188);

    let baseline = cardTop + 250 + 50;
    ctx.fillStyle = '#9a9a92';
    ctx.font = `26px ${FONT}`;
    ctx.textAlign = 'left';
    ctx.fillText('项目', 132, baseline);
    cols.forEach((col, index) => {
      ctx.textAlign = 'right';
      ctx.fillText(col.label, xs[index], baseline);
    });

    rowsShown.forEach((section) => {
      baseline += 56;
      ctx.font = `600 30px ${FONT}`;
      ctx.fillStyle = INK;
      ctx.textAlign = 'left';
      let text = section.name;
      if (ctx.measureText(text).width > (cols.length > 1 ? 430 : 560)) {
        const max = cols.length > 1 ? 430 : 560;
        while (text.length > 1 && ctx.measureText(`${text}…`).width > max) {
          text = text.slice(0, -1);
        }
        text = `${text}…`;
      }
      ctx.fillText(text, 132, baseline);
      cols.forEach((col, index) => {
        ctx.textAlign = 'right';
        ctx.font = `600 30px ${FONT}`;
        ctx.fillStyle = col.key === 'paid' ? PAY : col.key === 'hours' ? OT : INK;
        const cell =
          col.key === 'days'
            ? `${section.days} 天`
            : col.key === 'hours'
              ? `${section.hours} 小时`
              : col.key === 'paid'
                ? yuan(section.paid)
                : yuan(section.wage);
        ctx.fillText(cell, xs[index], baseline);
      });
    });

    if (sections.length > rowsShown.length) {
      baseline += 44;
      ctx.textAlign = 'left';
      ctx.font = `26px ${FONT}`;
      ctx.fillStyle = '#9a9a92';
      ctx.fillText(`…还有 ${sections.length - rowsShown.length} 个项目`, 132, baseline);
    }

    baseline += 62;
    ctx.textAlign = 'left';
    ctx.font = `700 30px ${FONT}`;
    ctx.fillStyle = INK;
    ctx.fillText('合计', 132, baseline);
    cols.forEach((col, index) => {
      ctx.textAlign = 'right';
      ctx.fillStyle = col.key === 'paid' ? PAY : col.key === 'hours' ? OT : GREEN;
      const cell =
        col.key === 'days'
          ? `${totals.days} 天`
          : col.key === 'hours'
            ? `${totals.hours} 小时`
            : col.key === 'paid'
              ? yuan(totals.paid)
              : yuan(totals.wage);
      ctx.fillText(cell, xs[index], baseline);
    });

    if (cols.length > 0 && (cols[0].key === 'wage' || cols[0].key === 'paid')) {
      baseline += 40;
      divider(baseline);
      baseline += 60;
      const summary = [];
      if (show.days) summary.push(`${totals.days} 工日`);
      if (show.overtime) summary.push(`加班 ${totals.hours} 小时`);
      ctx.textAlign = 'left';
      ctx.font = `28px ${FONT}`;
      ctx.fillStyle = MUTED;
      ctx.fillText(summary.join(' · ') || ' ', 132, baseline);
    }
  }

  ctx.fillStyle = '#9a9a92';
  ctx.font = `26px ${FONT}`;
  ctx.textAlign = 'center';
  ctx.fillText(`记工 · ${generatedAt}`, 540, 1250);

  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => (blob ? resolve(blob) : reject(new Error('toBlob failed'))),
      'image/png'
    );
  });
}

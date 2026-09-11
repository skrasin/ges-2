const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5

// ---- Дизайн-система ГЭС-2 (design-system/tokens.css) ----
const INK = '000000';
const PAPER = 'FFFFFF';
const MUTED = '9CA3AF';
const HINT = 'C0C1C0';
const HAIR = 'EEEEEE';
const FONT = 'Arial'; // fallback: Diagramatika проприетарная

const W = 13.333, H = 7.5;
const ML = 0.7, MR = 0.7;
const CW = W - ML - MR; // 11.933

const IMG = '/home/user/ges-2/docs/strategy/deck-images/';
const AR = 3 / 2; // все снимки заранее обрезаны ровно под 3:2

pres.defineSlideMaster({ title: 'GES2', background: { color: PAPER } });
const S = () => pres.addSlide({ masterName: 'GES2' });

// Заголовок слайда: Display - обычное начертание, трекинг +2%
function title(slide, text, opts = {}) {
  slide.addText(text, {
    x: ML, y: 0.5, w: CW, h: 0.8,
    fontFace: FONT, fontSize: 30, color: INK, charSpacing: 0.6,
    align: 'left', valign: 'top', margin: 0, isTextBox: true, ...opts,
  });
}

// Фотография: высота строго по ширине, без растяжения
function pic(slide, x, y, w, file) {
  const h = w / AR;
  slide.addImage({ path: IMG + file, x, y, w, h });
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { type: 'none' }, line: { color: INK, width: 0.5 },
  });
  return h;
}

// Плейсхолдер изображения
function photo(slide, x, y, w, caption) {
  const h = w / AR;
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h, fill: { color: HAIR }, line: { color: INK, width: 0.5 },
  });
  slide.addText(caption, {
    x: x + 0.2, y: y + h / 2 - 0.45, w: w - 0.4, h: 0.9,
    fontFace: FONT, fontSize: 11, color: MUTED, align: 'center', valign: 'middle',
    margin: 0, isTextBox: true,
  });
  return h;
}

function rule(slide, x, y, w) {
  slide.addShape(pres.ShapeType.line, { x, y, w, h: 0, line: { color: HAIR, width: 0.75 } });
}

function bullets(slide, items, opts = {}) {
  const runs = items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 10 },
  }));
  slide.addText(runs, {
    fontFace: FONT, fontSize: 14, color: INK, margin: 0, isTextBox: true,
    lineSpacingMultiple: 1.15, ...opts,
  });
}

// ============ 1. Титульный ============
{
  const s = S();
  s.addText('Дом культуры «ГЭС-2»', {
    x: ML, y: 0.6, w: CW, h: 0.3,
    fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, isTextBox: true,
  });
  s.addText('Предложение по улучшению\nнаправления сервиса\nи клиентского опыта ГЭС-2', {
    x: ML, y: 2.0, w: 10.5, h: 2.6,
    fontFace: FONT, fontSize: 38, color: INK, charSpacing: 0.76,
    lineSpacingMultiple: 1.1, margin: 0, isTextBox: true,
  });
  s.addText('«Аудитория доверяет ГЭС-2 открывать новое искусство»', {
    x: ML, y: 5.15, w: 9.5, h: 0.4,
    fontFace: FONT, fontSize: 16, color: INK, italic: true, margin: 0, isTextBox: true,
  });
  s.addText('Сентябрь 2026', {
    x: ML, y: 6.55, w: 8, h: 0.3,
    fontFace: FONT, fontSize: 11, color: MUTED, margin: 0, isTextBox: true,
  });
  s.addNotes('Титул. Оформление по дизайн-системе ГЭС-2: монохром, без скруглений и теней, вес на типографике. Шрифт Diagramatika проприетарный - верстка на fallback (Arial / Helvetica Neue).');
}

// ============ 2. Наше понимание задачи ============
{
  const s = S();
  title(s, 'Наше понимание задачи');
  const st = [
    'Доверие аудитории - то,\nради чего существует сервис',
    'Сервис не создаёт доверие,\nон его исполняет и конвертирует',
    'Периметр - весь путь, от поиска\nинформации до выхода из здания',
  ];
  let y = 1.95;
  st.forEach((t, i) => {
    s.addText(t, {
      x: ML, y, w: 6.2, h: 1.0,
      fontFace: FONT, fontSize: 18, color: INK, charSpacing: 0.36,
      lineSpacingMultiple: 1.15, margin: 0, isTextBox: true, valign: 'top',
    });
    if (i < st.length - 1) rule(s, ML, y + 1.18, 6.2);
    y += 1.45;
  });
  pic(s, 7.3, 2.3, 5.33, 'facade.jpg');
  s.addNotes('Три утверждения, задающие рамку. Фото фасада - узнаваемость за секунду.');
}

// ============ 3. Содержание ============
{
  const s = S();
  title(s, 'Содержание');
  const items = [
    'Горизонтальные связи',
    'Диагностика',
    'Обратная связь и метрики',
    'Персонал и гостеприимство',
    'Стандарты и документы',
    'Цифровые сервисы',
    'Площадка',
    'Монетизация и продукты',
  ];
  let y = 1.9;
  items.forEach((t, i) => {
    s.addText(String(i + 1).padStart(2, '0'), {
      x: ML, y, w: 0.6, h: 0.4,
      fontFace: FONT, fontSize: 13, color: MUTED, margin: 0, isTextBox: true,
    });
    s.addText(t, {
      x: ML + 0.7, y, w: 5.0, h: 0.4,
      fontFace: FONT, fontSize: 16, color: INK, charSpacing: 0.32, margin: 0, isTextBox: true,
    });
    rule(s, ML, y + 0.45, 5.7);
    y += 0.58;
  });
  pic(s, 6.9, 2.35, 5.73, 'prospekt.jpg');
  s.addNotes('Восемь направлений. Дальше каждое раскрывается отдельным слайдом.');
}

// ============ 4. Как выстраивается взаимодействие ============
{
  const s = S();
  title(s, 'Как выстраивается взаимодействие');
  pic(s, ML, 2.2, 5.2, 'svody.jpg');
  bullets(s, [
    'Регулярные совместные совещания с продюсерским и коммерческим блоками',
    'Ротация сотрудников бэк-офиса в зал',
    'Премиальный фонд для смежных подразделений в коммерческих проектах',
    'Внутренняя мотивация и делегирование ответственности',
    'Рабочая коммуникация в каналах, где люди действительно есть',
  ], { x: 6.45, y: 2.05, w: 6.18, h: 4.0, fontSize: 15 });
  s.addNotes('Механизмы, а не намерения: горизонталь строится процедурами и деньгами, не призывами.');
}

// ============ 5. Диагностика ============
{
  const s = S();
  title(s, 'Диагностика');
  const items = [
    'Пройти путь самостоятельно - цифровой и физический',
    'Профили посетителей и карты пути',
    'Service Blueprint: кто обеспечивает каждую точку',
    'Аудит площадки',
    'Интервью: фронт-линия и посетители',
  ];
  let y = 2.05;
  items.forEach((t, i) => {
    s.addText(String(i + 1), {
      x: ML, y: y - 0.04, w: 0.5, h: 0.5,
      fontFace: FONT, fontSize: 22, color: HINT, charSpacing: 0.44, margin: 0, isTextBox: true,
    });
    s.addText(t, {
      x: ML + 0.62, y, w: 5.5, h: 0.7,
      fontFace: FONT, fontSize: 15, color: INK, margin: 0, isTextBox: true,
      lineSpacingMultiple: 1.15, valign: 'top',
    });
    y += 0.82;
  });
  pic(s, 7.55, 2.4, 5.08, 'flow.jpg');
  s.addNotes('Порядок значим. Фотографию лучше заменить на собственный снимок реальной точки трения, сделанный при проходе пути.');
}

// ============ 6. Обратная связь и метрики ============
{
  const s = S();
  title(s, 'Обратная связь и метрики');
  const head = (t) => ({ text: t, options: { bold: true, fontSize: 14, color: INK } });
  const cell = (t, o = {}) => ({ text: t, options: { fontSize: 14, color: INK, ...o } });
  const rows = [
    [head('Что меряем'), head('Чем'), head('Как часто')],
    [cell('Удовлетворённость'), cell('CSI / CSAT'), cell('непрерывно', { color: MUTED })],
    [cell('Лёгкость взаимодействия'), cell('CES'), cell('непрерывно', { color: MUTED })],
    [cell('Лояльность'), cell('NPS'), cell('квартал', { color: MUTED })],
    [cell('Доверие'), cell('индекс доверия'), cell('год', { color: MUTED })],
    [cell('Удержание'), cell('повторные визиты'), cell('месяц', { color: MUTED })],
    [cell('Средний чек'), cell('per capita'), cell('месяц', { color: MUTED })],
  ];
  s.addTable(rows, {
    x: ML, y: 1.95, w: CW, colW: [5.2, 3.5, 3.233],
    rowH: 0.46, fontFace: FONT, valign: 'middle',
    border: [
      { type: 'none' }, { type: 'none' },
      { type: 'solid', color: HAIR, pt: 0.75 }, { type: 'none' },
    ],
    margin: [4, 0, 4, 0],
  });

  rule(s, ML, 5.45, CW);
  const extras = [
    'Журнал обращений и отзывы\nрегулярный разбор',
    'Тайный визит\nраз в год',
    'Бэклог запросов\nс оценкой стоимости',
  ];
  const ew = 3.71, eg = 0.4;
  extras.forEach((t, i) => {
    const parts = t.split('\n');
    s.addText([
      { text: parts[0], options: { bold: true, breakLine: true } },
      { text: parts[1], options: { color: MUTED } },
    ], {
      x: ML + i * (ew + eg), y: 5.7, w: ew, h: 0.9,
      fontFace: FONT, fontSize: 13, color: INK, margin: 0, isTextBox: true,
      lineSpacingMultiple: 1.2, valign: 'top',
    });
  });
  s.addNotes('Минимум показателей, по которым реально принимаются решения. Всё начинается с нулевого замера. Под таблицей - три постоянных источника обратной связи.');
}

// ============ 7. Персонал и гостеприимство ============
{
  const s = S();
  title(s, 'Персонал и гостеприимство');
  pic(s, ML, 1.95, 4.3, 'mediation.jpg');

  s.addText('Навык', {
    x: 5.55, y: 1.95, w: 7.08, h: 0.35,
    fontFace: FONT, fontSize: 16, color: INK, bold: true, margin: 0, isTextBox: true,
  });
  bullets(s, [
    'Гостеприимство как проверяемый навык: обучение, тестирование, внешние тренинги',
    'Актуальные скрипты и база знаний онлайн',
    'Регулярные воркшопы по программе',
    'Гибридная роль «сервис + контент»',
  ], { x: 5.55, y: 2.4, w: 7.08, h: 1.9, fontSize: 13 });

  s.addText('Условия', {
    x: 5.55, y: 4.45, w: 7.08, h: 0.35,
    fontFace: FONT, fontSize: 16, color: INK, bold: true, margin: 0, isTextBox: true,
  });
  bullets(s, [
    'Карьерная лестница, поощрения, оплата, графики',
    'Участие в общих собраниях, канал для идей с обязательным ответом',
    'Директор по клиентскому опыту',
  ], { x: 5.55, y: 4.9, w: 7.08, h: 1.5, fontSize: 13 });
  s.addNotes('Разделение намеренное: одним обучением дефицит гостеприимства не закрывается, условия работы влияют не меньше.');
}

// ============ 8. Стандарты и документы ============
{
  const s = S();
  title(s, 'Стандарты и документы');
  const docs = [
    'Стандарт\nобслуживания',
    'Публичное обещание\nпосетителю',
    'Хендбук сотрудника\nслужбы',
    'Регламент работы\nс обратной связью',
    'Матрица полномочий\nфронт-линии',
  ];
  const dw = 2.2, gap = 0.23;
  docs.forEach((d, i) => {
    const x = ML + i * (dw + gap);
    s.addShape(pres.ShapeType.rect, {
      x, y: 2.1, w: dw, h: 3.2,
      fill: { color: PAPER }, line: { color: INK, width: 0.75 },
    });
    s.addText(String(i + 1).padStart(2, '0'), {
      x: x + 0.22, y: 2.3, w: 1.0, h: 0.3,
      fontFace: FONT, fontSize: 12, color: MUTED, margin: 0, isTextBox: true,
    });
    s.addText(d, {
      x: x + 0.22, y: 3.9, w: dw - 0.44, h: 1.2,
      fontFace: FONT, fontSize: 14, color: INK, margin: 0, isTextBox: true,
      lineSpacingMultiple: 1.15, valign: 'bottom',
    });
  });
  s.addNotes('Пять артефактов, которые появляются на выходе. Порядок: сначала диагностика, потом документы.');
}

// ============ 9. Цифровые сервисы ============
{
  const s = S();
  title(s, 'Цифровые сервисы');
  bullets(s, [
    'UI/UX сайта и цифровые пути',
    'Мобильное приложение и веб-сервисы ГЭС-2 для посетителей',
    'Жизненный цикл систем, переход на микросервисную архитектуру',
  ], { x: ML, y: 2.05, w: 5.2, h: 2.2, fontSize: 15 });
  photo(s, 6.4, 2.2, 6.23, 'Скриншот ges-2.org');
  s.addNotes('Скриншот главной страницы сайта. Снимается самостоятельно - из этой среды внешние сайты недоступны.');
}

// ============ 10. Площадка ============
{
  const s = S();
  title(s, 'Площадка');
  bullets(s, [
    'Чат неполадок: увидел - сообщил - починили',
    'Улучшение точек питания',
    'SLA с внешними операторами: питание, клининг, паркинг, билетный сервис',
  ], { x: ML, y: 2.05, w: 5.5, h: 2.4, fontSize: 15 });
  pic(s, 6.9, 2.35, 5.73, 'platforms.jpg');
  s.addNotes('Качество на стыках с внешними операторами обеспечивается договором, а не просьбами.');
}

// ============ 11. Монетизация и продукты ============
{
  const s = S();
  title(s, 'Монетизация и продукты');

  s.addText('На существующем потоке', {
    x: ML, y: 1.95, w: 3.8, h: 0.35,
    fontFace: FONT, fontSize: 16, color: INK, bold: true, margin: 0, isTextBox: true,
  });
  bullets(s, [
    'Средний чек: магазин, питание, платные форматы',
    'Патронаж и членство для частных лиц',
  ], { x: ML, y: 2.4, w: 3.8, h: 1.8, fontSize: 13 });

  s.addText('Новые продукты', {
    x: 4.75, y: 1.95, w: 3.7, h: 0.35,
    fontFace: FONT, fontSize: 16, color: INK, bold: true, margin: 0, isTextBox: true,
  });
  bullets(s, [
    'Группа коммерческих сервисов',
    'Монетизация «Сводов»',
    'Инкубатор гипотез в MVP',
    'Услуги на внешнем рынке, арт-консалтинг',
    'Выставки корпоративных коллекций партнёров',
  ], { x: 4.75, y: 2.4, w: 3.7, h: 3.2, fontSize: 13 });

  pic(s, 8.7, 2.4, 3.93, 'books.jpg');
  s.addNotes('Платить за то, чтобы получить больше, укрепляет доверие. Платить за то, чтобы войти, испытывает его.');
}

pres.writeFile({ fileName: process.argv[2] || 'deck.pptx' }).then(f => console.log('OK:', f));

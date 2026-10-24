require('dotenv').config({ path: require('path').join(__dirname, '.env') });

const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const multer = require('multer');
const QR = require('qrcode');
const Excel = require('exceljs');
const PDF = require('pdfkit');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const S = process.env.JWT_SECRET;

if (!S) {
  console.error('ERROR: JWT_SECRET is missing');
  process.exit(1);
}
const CLIENT = process.env.CLIENT_URL;
const PORT = process.env.PORT || 5000;
const MONGO_URI = process.env.MONGO_URI;

if (!MONGO_URI) {
  console.error('ERROR: MONGO_URI is missing. Check server/.env');
  process.exit(1);
}

fs.mkdirSync(path.join(__dirname, 'uploads'), { recursive: true });

const up = multer({
  storage: multer.diskStorage({
    destination: path.join(__dirname, 'uploads'),
    filename: (q, f, cb) =>
      cb(null, Date.now() + '-' + f.originalname.replace(/[^\w.-]/g, '_'))
  })
});

const Id = mongoose.Schema.Types.ObjectId;
const M = (n, s) => mongoose.model(n, new mongoose.Schema(s, { timestamps: true }));

const User = M('User', {
  name: String,
  email: { type: String, unique: true },
  hash: String,
  role: { type: String, default: 'student' }
});

const Event = M('Event', {
  title: String,
  description: String,
  category: String,
  date: String,
  time: String,
  venue: String,
  banner: String,
  seats: Number,
  deadline: String,
  coordinator: String,
  completed: { type: Boolean, default: false }
});

const Reg = M('Reg', {
  user: { type: Id, ref: 'User' },
  event: { type: Id, ref: 'Event' },
  ticket: String,
  attended: { type: Boolean, default: false },
  attendedAt: Date
});

const Cert = M('Cert', {
  user: { type: Id, ref: 'User' },
  event: { type: Id, ref: 'Event' },
  type: String,
  cid: String
});

const Gallery = M('Gallery', {
  title: String,
  category: { type: String, default: 'Campus' },
  image: String,
  description: String
});

const Setting = M('Setting', {
  collegeName: String,
  logo: String,
  hero: String,
  signature: String,
  signatory: String,
  designation: String,
  address: String,
  contactName: String,
  contactEmail: String,
  contactPhone: String
});

const cfg = async () => {
  let c = await Setting.findOne();

  if (!c) {
    c = await Setting.create({
      collegeName: 'Adikavi Nannaya University',
      logo: '/assets/college-logo.png',
      hero: '/assets/college-gate.png',
      signature: '/assets/signature.png',
      signatory: 'Registrar',
      designation: 'Adikavi Nannaya University',
      address: 'Rajamahendravaram, Andhra Pradesh',
      contactName: 'Event Coordinator Cell',
      contactEmail: 'events@aknu.edu.in',
      contactPhone: '+91 883 2566004 / 2566009'
    });
  }

  if (!c.signature) {
    c.signature = '/assets/signature.png';
    await c.save();
  }

  if (c.address === undefined) {
    c.address = 'Rajamahendravaram, Andhra Pradesh';
    await c.save();
  }
  if (!c.contactName) {
    c.contactName = 'Event Coordinator Cell';
    await c.save();
  }
  if (!c.contactEmail) {
    c.contactEmail = 'events@aknu.edu.in';
    await c.save();
  }
  if (!c.contactPhone) {
    c.contactPhone = '+91 883 2566004 / 2566009';
    await c.save();
  }

  return c;
};

const auth = role => (q, s, n) => {
  try {
    q.u = jwt.verify((q.headers.authorization || '').split(' ')[1], S);
    if (role && q.u.role !== role) throw new Error('Unauthorized');
    n();
  } catch {
    s.status(401).json({ error: 'Please log in again' });
  }
};

const wrap = f => (q, s, n) =>
  Promise.resolve(f(q, s, n)).catch(e => {
    if (!s.headersSent) s.status(400).json({ error: e.message });
  });

const sign = u => ({
  token: jwt.sign({ id: u._id, role: u.role }, S, { expiresIn: '7d' }),
  user: { id: u._id, name: u.name, role: u.role }
});

const newTicket = async () => {
  for (; ;) {
    const t = 'AKNU-' + crypto.randomBytes(4).toString('hex').toUpperCase().slice(0, 6);
    if (!await Reg.findOne({ ticket: t })) return t;
  }
};

const fp = u =>
  u && (u.startsWith('/uploads') || u.startsWith('/assets'))
    ? path.join(__dirname, u)
    : null;

const app = express();

app.use(cors({
  origin: (origin, callback) => {
    if (!origin) return callback(null, true);
    if (
      !CLIENT ||
      CLIENT === '*' ||
      origin === CLIENT ||
      origin.endsWith('.vercel.app') ||
      origin.includes('localhost') ||
      origin.includes('127.0.0.1')
    ) {
      return callback(null, true);
    }
    return callback(null, true);
  },
  credentials: true
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.use('/uploads', cors(), express.static(path.join(__dirname, 'uploads'), {
  setHeaders: (res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cross-Origin-Resource-Policy', 'cross-origin');
  }
}));
app.use('/assets', cors(), express.static(path.join(__dirname, 'assets'), {
  setHeaders: (res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cross-Origin-Resource-Policy', 'cross-origin');
  }
}));

app.get('/api/health', (q, s) => {
  s.json({ status: 'ok', database: mongoose.connection.readyState === 1 ? 'connected' : 'disconnected' });
});

app.get('/api/settings', wrap(async (q, s) => s.json(await cfg())));

app.put(
  '/api/settings',
  auth('admin'),
  up.fields([{ name: 'logo' }, { name: 'hero' }, { name: 'signature' }]),
  wrap(async (q, s) => {
    const c = await cfg();

    for (const k of ['collegeName', 'address', 'signatory', 'designation', 'contactName', 'contactEmail', 'contactPhone']) {
      if (q.body[k] !== undefined) c[k] = q.body[k];
    }

    for (const k of ['logo', 'hero', 'signature']) {
      if (q.files?.[k]) c[k] = '/uploads/' + q.files[k][0].filename;
    }

    await c.save();
    s.json(c);
  })
);

app.post('/api/auth/register', wrap(async (q, s) => {
  const { name, email, password } = q.body;

  if (!name || !email || !password || password.length < 6) {
    throw Error('Enter name, email and a password of 6+ characters');
  }

  const normalizedEmail = email.toLowerCase();

  if (await User.findOne({ email: normalizedEmail })) {
    throw Error('This email is already registered');
  }

  const user = await User.create({
    name,
    email: normalizedEmail,
    hash: await bcrypt.hash(password, 10)
  });

  s.json(sign(user));
}));

const lg = role => wrap(async (q, s) => {
  const identifier = String(q.body.email || q.body.id || '').toLowerCase();

  const u = await User.findOne({
    email: identifier,
    role
  });

  if (!u || !await bcrypt.compare(q.body.password || '', u.hash)) {
    throw Error('Wrong ID/email or password');
  }

  s.json(sign(u));
});

app.post('/api/auth/login', lg('student'));
app.post('/api/auth/admin', lg('admin'));

app.put('/api/auth/profile', auth(), wrap(async (q, s) => {
  const { name, currentPassword, newPassword } = q.body;
  const u = await User.findById(q.u.id);
  if (!u) throw Error('User not found');

  if (newPassword) {
    if (newPassword.length < 6) {
      throw Error('New password must be at least 6 characters');
    }
    if (!currentPassword) {
      throw Error('Current password is required to change password');
    }
    const match = await bcrypt.compare(currentPassword, u.hash);
    if (!match) {
      throw Error('Current password is incorrect');
    }
    u.hash = await bcrypt.hash(newPassword, 10);
  }

  if (name && name.trim()) {
    u.name = name.trim();
  }

  await u.save();

  s.json({
    success: true,
    user: { id: u._id, name: u.name, role: u.role },
    message: 'Profile updated successfully'
  });
}));

app.get('/api/events', wrap(async (q, s) => {
  const ev = await Event.find().sort('date').lean();

  for (const e of ev) {
    e.taken = await Reg.countDocuments({ event: e._id });
  }

  s.json(ev);
}));

const defaultGallery = [
  {
    title: 'University Grand Entrance',
    category: 'Campus',
    image: '/assets/college-gate.png',
    description: 'The magnificent main gate and lush academic campus of Adikavi Nannaya University.'
  },
  {
    title: 'Annual Cultural Fest - Youth Tarang',
    category: 'Cultural',
    image: 'https://images.unsplash.com/photo-1514525253161-7a46d19cd819?w=800&auto=format&fit=crop&q=80',
    description: 'Students celebrating arts, traditional classical dance, and folk music on stage.'
  },
  {
    title: 'Hackathon & Tech Wiz Competition',
    category: 'Technical',
    image: 'https://images.unsplash.com/photo-1531482615713-2afd69097998?w=800&auto=format&fit=crop&q=80',
    description: 'Inter-college competitive programming, robotics exhibitions, and innovative tech projects.'
  },
  {
    title: 'Annual Sports Day & Athletics Meet',
    category: 'Sports',
    image: 'https://images.unsplash.com/photo-1461896836934-ffe607ba8211?w=800&auto=format&fit=crop&q=80',
    description: 'Track and field events, football, cricket tournaments, and athletic championship ceremonies.'
  },
  {
    title: 'Grand University Auditorium',
    category: 'Campus',
    image: 'https://images.unsplash.com/photo-1540575467063-178a50c2df87?w=800&auto=format&fit=crop&q=80',
    description: 'Air-conditioned modern auditorium hosting national conferences, workshops, and guest summits.'
  },
  {
    title: 'Convocation & Degree Felicitations',
    category: 'Academic',
    image: 'https://images.unsplash.com/photo-1523050854058-8df90110c9f1?w=800&auto=format&fit=crop&q=80',
    description: 'Honoring outstanding graduates, medal winners, and celebrating university milestone achievements.'
  }
];

app.get('/api/gallery', wrap(async (q, s) => {
  let items = await Gallery.find().sort('createdAt').lean();
  if (!items.length) {
    items = await Gallery.insertMany(defaultGallery);
  }
  s.json(items);
}));

app.post('/api/gallery', auth('admin'), up.single('image'), wrap(async (q, s) => {
  const { title, category, description, imageUrl } = q.body;
  let image = '';
  if (q.file) {
    image = '/uploads/' + q.file.filename;
  } else if (imageUrl && imageUrl.trim()) {
    image = imageUrl.trim();
  }
  if (!title) throw Error('Photo title is required');
  const item = await Gallery.create({
    title: title.trim(),
    category: category || 'Campus',
    description: description || '',
    image
  });
  s.json(item);
}));

app.delete('/api/gallery/:id', auth('admin'), wrap(async (q, s) => {
  const item = await Gallery.findByIdAndDelete(q.params.id);
  if (!item) throw Error('Gallery photo not found');
  s.json({ ok: 1 });
}));

app.post(
  '/api/events',
  auth('admin'),
  up.single('banner'),
  wrap(async (q, s) => {
    s.json(await Event.create({
      ...q.body,
      banner: q.file ? '/uploads/' + q.file.filename : ''
    }));
  })
);

app.delete('/api/events/:id', auth('admin'), wrap(async (q, s) => {
  await Event.findByIdAndDelete(q.params.id);
  await Reg.deleteMany({ event: q.params.id });
  await Cert.deleteMany({ event: q.params.id });
  s.json({ ok: 1 });
}));

app.patch('/api/events/:id/toggle', auth('admin'), wrap(async (q, s) => {
  const e = await Event.findById(q.params.id);
  if (!e) throw Error('Event not found');
  if (typeof q.body.completed === 'boolean') {
    e.completed = q.body.completed;
  } else if (typeof q.body.active === 'boolean') {
    e.completed = !q.body.active;
  } else {
    e.completed = !e.completed;
  }
  await e.save();
  s.json({ ok: 1, completed: e.completed, active: !e.completed, event: e });
}));

app.put('/api/events/:id', auth('admin'), wrap(async (q, s) => {
  const e = await Event.findById(q.params.id);
  if (!e) throw Error('Event not found');
  ['title', 'description', 'category', 'date', 'time', 'venue', 'seats', 'deadline', 'coordinator'].forEach(f => {
    if (q.body[f] !== undefined) e[f] = q.body[f];
  });
  if (typeof q.body.completed === 'boolean') e.completed = q.body.completed;
  if (typeof q.body.active === 'boolean') e.completed = !q.body.active;
  await e.save();
  s.json({ ok: 1, completed: e.completed, active: !e.completed, event: e });
}));

app.post('/api/events/:id/register', auth('student'), wrap(async (q, s) => {
  const e = await Event.findById(q.params.id);

  if (!e) throw Error('Event not found');
  if (e.completed) throw Error('This event is completed');

  if (e.deadline && new Date(e.deadline + 'T23:59:59') < new Date()) {
    throw Error('Registration deadline has passed');
  }

  if (await Reg.findOne({ user: q.u.id, event: e._id })) {
    throw Error('You are already registered');
  }

  if (e.seats && await Reg.countDocuments({ event: e._id }) >= e.seats) {
    throw Error('No seats left');
  }

  await Reg.create({
    user: q.u.id,
    event: e._id,
    ticket: await newTicket()
  });

  s.json({ ok: 1 });
}));

app.get('/api/my/regs', auth('student'), wrap(async (q, s) => {
  const rs = await Reg.find({ user: q.u.id })
    .populate('event')
    .populate('user', 'name')
    .lean();

  for (const r of rs) {
    if (!r.ticket) {
      r.ticket = await newTicket();
      await Reg.findByIdAndUpdate(r._id, { ticket: r.ticket });
    }

    r.qr = await QR.toDataURL(r.ticket, { width: 400, margin: 4 });
  }

  s.json(rs);
}));

app.get('/api/my/certs', auth('student'), wrap(async (q, s) => {
  s.json(
    await Cert.find({ user: q.u.id })
      .populate('event', 'title date')
      .lean()
  );
}));

app.get('/api/events/:id/regs', auth('admin'), wrap(async (q, s) => {
  const filter = q.params.id === 'all' ? {} : { event: q.params.id };
  s.json(
    await Reg.find(filter)
      .populate('user', 'name email')
      .populate('event', 'title date venue')
      .sort('-createdAt')
      .lean()
  );
}));

app.post('/api/regs/:id/checkin', auth('admin'), wrap(async (q, s) => {
  const r = await Reg.findById(q.params.id);

  if (!r) throw Error('Registration not found');
  if (r.attended) throw Error('Already checked in');

  r.attended = true;
  r.attendedAt = new Date();
  await r.save();

  s.json({ ok: 1 });
}));

app.put('/api/regs/:id', auth('admin'), wrap(async (q, s) => {
  const r = await Reg.findById(q.params.id);
  if (!r) throw Error('Registration not found');

  const { name, email, attended, eventId } = q.body;

  if (attended !== undefined) {
    r.attended = Boolean(attended);
    if (r.attended && !r.attendedAt) r.attendedAt = new Date();
    if (!r.attended) r.attendedAt = null;
  }

  if (eventId && mongoose.Types.ObjectId.isValid(eventId)) {
    r.event = eventId;
  }

  await r.save();

  if (r.user && (name || email)) {
    const u = await User.findById(r.user);
    if (u) {
      if (name) u.name = name.trim();
      if (email) u.email = email.trim().toLowerCase();
      await u.save();
    }
  }

  const updated = await Reg.findById(r._id)
    .populate('user', 'name email')
    .populate('event', 'title date venue')
    .lean();

  s.json({ ok: 1, registration: updated });
}));

app.delete('/api/regs/:id', auth('admin'), wrap(async (q, s) => {
  const r = await Reg.findByIdAndDelete(q.params.id);
  if (!r) throw Error('Registration not found');
  s.json({ ok: 1 });
}));

app.post('/api/attendance/scan', auth('admin'), wrap(async (q, s) => {
  const g = await Reg.findOne({
    ticket: String(q.body.token || '').trim().toUpperCase()
  }).populate('user event');

  if (!g) throw Error('Ticket not found');
  if (g.attended) throw Error(g.user.name + ' is already checked in');

  g.attended = true;
  g.attendedAt = new Date();
  await g.save();

  s.json({ name: g.user.name, event: g.event.title });
}));

app.get('/api/stats', auth('admin'), wrap(async (q, s) => {
  s.json({
    events: await Event.countDocuments(),
    regs: await Reg.countDocuments(),
    attended: await Reg.countDocuments({ attended: true }),
    certs: await Cert.countDocuments(),
    recent: await Reg.find()
      .sort('-createdAt')
      .limit(8)
      .populate('user', 'name')
      .populate('event', 'title')
      .lean()
  });
}));

app.get('/api/events/:id/export/:fmt', auth('admin'), wrap(async (q, s) => {
  const isAll = q.params.id === 'all';
  let e = null;
  let rs = [];

  if (isAll) {
    e = { title: 'All Registered Events', date: 'Consolidated', venue: 'All Venues' };
    rs = await Reg.find()
      .populate('user', 'name email')
      .populate('event', 'title date venue')
      .sort('-createdAt')
      .lean();
  } else {
    e = await Event.findById(q.params.id);
    if (!e) throw Error('Event not found');
    rs = await Reg.find({ event: e._id })
      .populate('user', 'name email')
      .populate('event', 'title date venue')
      .sort('-createdAt')
      .lean();
  }

  const st = await cfg();

  const rows = rs.map((r, i) => [
    i + 1,
    r.user?.name || '',
    r.user?.email || '',
    r.ticket || '-',
    new Date(r.createdAt),
    r.attended ? 'Present' : 'Absent'
  ]);

  const H = ['#', 'Name', 'Email', 'Ticket ID', 'Registered on', 'Attendance'];
  const present = rows.filter(r => r[5] === 'Present').length;

  if (q.params.fmt === 'excel') {
    const w = new Excel.Workbook();
    const x = w.addWorksheet('Registrations');
    const bd = { style: 'thin', color: { argb: 'FFC9D1E6' } };
    const B = { top: bd, left: bd, bottom: bd, right: bd };

    x.mergeCells('A1:F1');
    x.getCell('A1').value = st.collegeName + ' - ' + e.title;
    x.getCell('A1').font = {
      bold: true,
      size: 15,
      color: { argb: 'FF0B1D3A' }
    };

    x.mergeCells('A2:F2');
    x.getCell('A2').value =
      `${e.date} | ${e.venue || ''} | Registered: ${rows.length} | Present: ${present}`;

    const hr = x.addRow(H);

    hr.eachCell(c => {
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0B1D3A' } };
      c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      c.border = B;
    });

    rows.forEach((r, i) => {
      const row = x.addRow(r);
      row.getCell(5).numFmt = 'dd-mmm-yyyy hh:mm';

      row.eachCell(c => {
        c.border = B;

        if (i % 2) {
          c.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF2F5FB' }
          };
        }
      });
    });

    [6, 26, 32, 16, 22, 14].forEach((wd, i) => {
      x.getColumn(i + 1).width = wd;
    });

    x.views = [{ state: 'frozen', ySplit: 3 }];
    x.autoFilter = { from: 'A3', to: 'F3' };

    s.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    s.setHeader('Content-Disposition', 'attachment; filename="registrations.xlsx"');

    await w.xlsx.write(s);
    return s.end();
  }

  if (q.params.fmt !== 'pdf') {
    throw Error('Unsupported export format');
  }

  const d = new PDF({ size: 'A4', margin: 40 });
  const C = [
    ['#', 28],
    ['Name', 118],
    ['Email', 150],
    ['Ticket ID', 72],
    ['Registered', 78],
    ['Status', 69]
  ];

  const lg = fp(st.logo);
  s.type('pdf');
  d.pipe(s);

  if (lg && fs.existsSync(lg)) d.image(lg, 40, 32, { height: 48 });

  d.fillColor('#0b1d3a')
    .font('Helvetica-Bold')
    .fontSize(16)
    .text(st.collegeName, 100, 36, { width: 415 })
    .font('Helvetica')
    .fontSize(11)
    .text(e.title + ' | ' + e.date + ' | ' + (e.venue || ''), 100, 58, { width: 415 })
    .fontSize(9)
    .fillColor('#5d6785')
    .text(`Registered: ${rows.length}   Present: ${present}   Absent: ${rows.length - present}`, 100, 74, { width: 415 });

  const line = (y, v, hdr, i = 0) => {
    let x = 40;

    C.forEach(([, w], k) => {
      d.rect(x, y, w, 22)
        .fillAndStroke(hdr ? '#0b1d3a' : i % 2 ? '#f2f5fb' : '#ffffff', '#c9d1e6');

      d.fillColor(hdr ? '#ffffff' : '#0b1d3a')
        .font(hdr ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(9)
        .text(String(v[k]), x + 4, y + 7, {
          width: w - 8,
          height: 12,
          ellipsis: true,
          lineBreak: false
        });

      x += w;
    });
  };

  let y = 100;
  line(y, C.map(c => c[0]), true);
  y += 22;

  rows.forEach((r, i) => {
    if (y > 760) {
      d.addPage();
      y = 40;
      line(y, C.map(c => c[0]), true);
      y += 22;
    }

    line(y, [
      r[0],
      r[1],
      r[2],
      r[3],
      r[4].toLocaleDateString('en-IN'),
      r[5]
    ], false, i);

    y += 22;
  });

  d.end();
}));

app.post('/api/events/:id/certs', auth('admin'), wrap(async (q, s) => {
  const rs = await Reg.find({ event: q.params.id }).lean();
  const ok = ['Participation', 'Winner 1st', 'Winner 2nd', 'Winner 3rd'];

  await Cert.deleteMany({ event: q.params.id });

  for (const it of q.body.items || []) {
    const r = rs.find(x => String(x.user) === it.user);

    if (!r || !ok.includes(it.type)) continue;
    if (it.type === 'Participation' && !r.attended) continue;

    await Cert.create({
      user: it.user,
      event: q.params.id,
      type: it.type,
      cid: crypto.randomBytes(8).toString('hex')
    });
  }

  await Event.findByIdAndUpdate(q.params.id, { completed: true });

  s.json({ count: await Cert.countDocuments({ event: q.params.id }) });
}));

app.get('/api/verify/:cid', wrap(async (q, s) => {
  const c = await Cert.findOne({ cid: q.params.cid })
    .populate('user', 'name')
    .populate('event', 'title date');

  if (!c) throw Error('Certificate not found');

  s.json({
    name: c.user.name,
    event: c.event.title,
    date: c.event.date,
    type: c.type
  });
}));

const CERT_THEMES = {
  'Winner 1st': {
    badge: '1ST',
    badgeText: '1ST PLACE WINNER',
    medalRibbon: 'GOLD AWARD',
    title: 'CERTIFICATE OF EXCELLENCE',
    subTitle: 'FOR OUTSTANDING MERIT & ACHIEVEMENT',
    awardedText: 'This Certificate is Proudly Awarded to',
    primaryColor: '#8a6208',
    primaryDark: '#5c4103',
    accentColor: '#d4af37',
    accentLight: '#fffdf6',
    borderOuter: '#8a6208',
    borderInner: '#d4af37',
    ribbonBg: '#8a6208',
    ribbonText: '#ffffff',
    nameColor: '#3d2b02',
    sealBg: '#fef3c7',
    sealBorder: '#b48512',
    sealRibbon: '#b48512',
    sealText: '#5c4103'
  },
  'Winner 2nd': {
    badge: '2ND',
    badgeText: '2ND PLACE - RUNNER UP',
    medalRibbon: 'SILVER AWARD',
    title: 'CERTIFICATE OF MERIT',
    subTitle: 'FOR DISTINGUISHED PERFORMANCE',
    awardedText: 'This Certificate is Proudly Awarded to',
    primaryColor: '#334155',
    primaryDark: '#0f172a',
    accentColor: '#64748b',
    accentLight: '#f8fafc',
    borderOuter: '#1e293b',
    borderInner: '#94a3b8',
    ribbonBg: '#334155',
    ribbonText: '#ffffff',
    nameColor: '#0f172a',
    sealBg: '#e2e8f0',
    sealBorder: '#475569',
    sealRibbon: '#475569',
    sealText: '#0f172a'
  },
  'Winner 3rd': {
    badge: '3RD',
    badgeText: '3RD PLACE',
    medalRibbon: 'BRONZE AWARD',
    title: 'CERTIFICATE OF COMMENDATION',
    subTitle: 'FOR PRAISEWORTHY ACCOMPLISHMENT',
    awardedText: 'This Certificate is Proudly Awarded to',
    primaryColor: '#7c2d12',
    primaryDark: '#431407',
    accentColor: '#c2410c',
    accentLight: '#fffaf5',
    borderOuter: '#7c2d12',
    borderInner: '#ea580c',
    ribbonBg: '#7c2d12',
    ribbonText: '#ffffff',
    nameColor: '#431407',
    sealBg: '#ffedd5',
    sealBorder: '#c2410c',
    sealRibbon: '#9a3412',
    sealText: '#431407'
  },
  'Participation': {
    badge: 'PART',
    badgeText: 'PARTICIPATION',
    medalRibbon: 'OFFICIAL MERIT',
    title: 'CERTIFICATE OF PARTICIPATION',
    subTitle: 'FOR COMMENDABLE DEDICATION & ENTHUSIASM',
    awardedText: 'This Certificate is Proudly Presented to',
    primaryColor: '#0b2545',
    primaryDark: '#05192d',
    accentColor: '#1d4ed8',
    accentLight: '#f6f9fe',
    borderOuter: '#0b2545',
    borderInner: '#3b82f6',
    ribbonBg: '#0b2545',
    ribbonText: '#ffffff',
    nameColor: '#0b2545',
    sealBg: '#dbeafe',
    sealBorder: '#1d4ed8',
    sealRibbon: '#1e40af',
    sealText: '#1e3a8a'
  }
};

app.get('/api/cert/:cid.pdf', wrap(async (q, s) => {
  const c = await Cert.findOne({ cid: q.params.cid }).populate('user event');

  if (!c) throw Error('Certificate not found');

  const st = await cfg();
  const theme = CERT_THEMES[c.type] || CERT_THEMES['Participation'];
  const W = 841.89;
  const H = 595.28;
  const d = new PDF({ size: 'A4', layout: 'landscape', margin: 0 });

  const ex = p => p && fs.existsSync(p) ? p : null;
  const hero = ex(fp(st.hero)) || ex(path.join(__dirname, 'assets', 'college-gate.png'));
  const lg = ex(fp(st.logo)) || ex(path.join(__dirname, 'assets', 'college-logo.png'));
  const sg = ex(fp(st.signature)) || ex(path.join(__dirname, 'assets', 'signature.png'));

  s.type('pdf');
  d.pipe(s);

  // 1. Background fill
  d.rect(0, 0, W, H).fill(theme.accentLight);

  // 2. Subtle campus watermark
  if (hero) {
    d.save()
      .opacity(0.04)
      .image(hero, (W - 550) / 2, (H - 350) / 2, { width: 550 })
      .restore();
  }

  // 3. Ornate Double Borders & Corner Accents
  d.rect(20, 20, W - 40, H - 40).lineWidth(3.5).stroke(theme.borderOuter);
  d.rect(25, 25, W - 50, H - 50).lineWidth(0.8).stroke(theme.borderInner);
  d.rect(30, 30, W - 60, H - 60).lineWidth(1.5).stroke(theme.borderOuter);

  function drawCorner(cx, cy) {
    d.save();
    d.rect(cx - 5, cy - 5, 10, 10).lineWidth(0.8).stroke(theme.borderInner);
    d.circle(cx, cy, 2.5).fill(theme.borderOuter);
    d.restore();
  }
  drawCorner(30, 30);
  drawCorner(W - 30, 30);
  drawCorner(30, H - 30);
  drawCorner(W - 30, H - 30);

  // 4. University Header
  if (lg) {
    d.image(lg, W / 2 - 26, 40, { width: 52, height: 52 });
  }

  d.font('Helvetica-Bold')
    .fontSize(21)
    .fillColor(theme.primaryDark)
    .text((st.collegeName || 'ADIKAVI NANNAYA UNIVERSITY').toUpperCase(), 0, 96, {
      align: 'center',
      width: W,
      characterSpacing: 1.5
    });

  d.font('Helvetica')
    .fontSize(9.5)
    .fillColor('#64748b')
    .text(st.address || 'Rajamahendravaram, Andhra Pradesh - 533296', 0, 120, {
      align: 'center',
      width: W
    });

  // Divider flourish
  d.moveTo(W / 2 - 160, 137).lineTo(W / 2 - 16, 137).lineWidth(0.8).stroke(theme.borderInner);
  d.moveTo(W / 2 + 16, 137).lineTo(W / 2 + 160, 137).lineWidth(0.8).stroke(theme.borderInner);
  d.save().translate(W / 2, 137).rotate(45).rect(-3, -3, 6, 6).fill(theme.borderOuter).restore();

  // 5. Certificate Title & Distinction Badge
  d.font('Helvetica-Bold')
    .fontSize(21)
    .fillColor(theme.primaryColor)
    .text(theme.title, 0, 150, { align: 'center', width: W, characterSpacing: 3.5 });

  const pillW = 240;
  d.roundedRect(W / 2 - pillW / 2, 180, pillW, 20, 10).fillAndStroke(theme.ribbonBg, theme.borderOuter);
  d.font('Helvetica-Bold')
    .fontSize(9)
    .fillColor(theme.ribbonText)
    .text(`•  ${theme.badgeText}  •`, 0, 185, { align: 'center', width: W, characterSpacing: 2 });

  // 6. Presentation Line
  d.font('Helvetica-Oblique')
    .fontSize(12)
    .fillColor('#555555')
    .text(theme.awardedText, 0, 218, { align: 'center', width: W });

  // 7. Student Name
  const rawName = c.user?.name || 'Student';
  const studentName = rawName.replace(/\b\w/g, l => l.toUpperCase());
  d.font('Helvetica-Bold')
    .fontSize(32)
    .fillColor(theme.nameColor)
    .text(studentName, 0, 240, { align: 'center', width: W });

  const uW = Math.min(380, Math.max(240, studentName.length * 15));
  d.moveTo(W / 2 - uW / 2, 280).lineTo(W / 2 + uW / 2, 280).lineWidth(1.2).stroke(theme.accentColor);
  d.save().translate(W / 2, 280).rotate(45).rect(-3, -3, 6, 6).fill(theme.primaryColor).restore();

  // 8. Citation Paragraph
  const eventTitle = c.event?.title || 'University Event';
  const eventDate = c.event?.date || '';
  const eventVenue = c.event?.venue || 'Campus Auditorium';
  const collegeName = st.collegeName || 'Adikavi Nannaya University';

  let citationText = '';
  if (c.type === 'Winner 1st') {
    citationText = `has accomplished exceptional performance, securing 1st Place (Winner) in the university event ${eventTitle} conducted on ${eventDate} at ${eventVenue}, organized by ${collegeName}.`;
  } else if (c.type === 'Winner 2nd') {
    citationText = `has accomplished outstanding performance, securing 2nd Place (Runner-Up) in the university event ${eventTitle} conducted on ${eventDate} at ${eventVenue}, organized by ${collegeName}.`;
  } else if (c.type === 'Winner 3rd') {
    citationText = `has accomplished commendable performance, securing 3rd Place in the university event ${eventTitle} conducted on ${eventDate} at ${eventVenue}, organized by ${collegeName}.`;
  } else {
    citationText = `has actively participated and contributed with excellence in the university event ${eventTitle} conducted on ${eventDate} at ${eventVenue}, organized by ${collegeName}.`;
  }

  d.font('Helvetica')
    .fontSize(13)
    .fillColor('#222222')
    .text(citationText, (W - 620) / 2, 300, { align: 'center', width: 620, lineGap: 5 });

  // 9. Official Embossed Rosette Seal (Bottom Center)
  const sealX = W / 2;
  const sealY = 445;

  d.polygon([sealX - 14, sealY + 22], [sealX - 24, sealY + 60], [sealX - 13, sealY + 52], [sealX - 2, sealY + 60], [sealX - 4, sealY + 22]).fill(theme.sealRibbon);
  d.polygon([sealX + 4, sealY + 22], [sealX + 2, sealY + 60], [sealX + 13, sealY + 52], [sealX + 24, sealY + 60], [sealX + 14, sealY + 22]).fill(theme.sealRibbon);

  d.circle(sealX, sealY, 34).lineWidth(2).fillAndStroke(theme.sealBorder, theme.borderOuter);
  d.circle(sealX, sealY, 30).fill(theme.sealBg);
  d.save().circle(sealX, sealY, 26).lineWidth(1).dash(2, { space: 2 }).stroke(theme.sealBorder).restore();

  d.font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(theme.sealText)
    .text(theme.badge, sealX - 30, sealY - 11, { width: 60, align: 'center' });

  d.font('Helvetica-Bold')
    .fontSize(6.5)
    .fillColor(theme.sealText)
    .text(theme.medalRibbon, sealX - 30, sealY + 5, { width: 60, align: 'center', characterSpacing: 0.8 });

  // 10. Security QR Code (Bottom Left)
  const qrBuffer = await QR.toBuffer(CLIENT + '/verify/' + c.cid, { margin: 1, width: 220 });

  d.roundedRect(56, 420, 84, 84, 8).fillAndStroke('#ffffff', theme.borderInner);
  d.image(qrBuffer, 60, 424, { width: 76 });

  d.font('Helvetica-Bold')
    .fontSize(7.5)
    .fillColor(theme.primaryDark)
    .text('VERIFIED AUTHENTIC', 46, 510, { width: 104, align: 'center' });

  d.font('Helvetica')
    .fontSize(7)
    .fillColor('#64748b')
    .text('ID: ' + c.cid, 44, 521, { width: 108, align: 'center' });

  // 11. Official Signature (Bottom Right)
  if (sg) {
    d.image(sg, W - 225, 422, { width: 130, height: 48, fit: [130, 48], align: 'center' });
  }

  d.moveTo(W - 240, 478).lineTo(W - 80, 478).lineWidth(1.2).stroke(theme.primaryDark);

  d.font('Helvetica-Bold')
    .fontSize(10.5)
    .fillColor(theme.primaryDark)
    .text(st.signatory || 'Registrar', W - 240, 485, { width: 160, align: 'center' });

  d.font('Helvetica')
    .fontSize(8.5)
    .fillColor('#64748b')
    .text(st.designation || st.collegeName, W - 240, 499, { width: 160, align: 'center' });

  d.font('Helvetica-Oblique')
    .fontSize(7.5)
    .fillColor('#94a3b8')
    .text('Authorized Signatory', W - 240, 513, { width: 160, align: 'center' });

  d.end();
}));

async function startServer() {
  try {
    console.log('Connecting to MongoDB...');

    await mongoose.connect(MONGO_URI, {
      serverSelectionTimeoutMS: 15000
    });

    console.log('MongoDB connected successfully');

    const id = (process.env.ADMIN_ID || 'admin').toLowerCase();

    if (!await User.findOne({ role: 'admin' })) {
      await User.create({
        name: 'Administrator',
        email: id,
        role: 'admin',
        hash: await bcrypt.hash(process.env.ADMIN_PASSWORD || 'admin123', 10)
      });

      console.log('Default admin account created');
    }

    if (!await Event.countDocuments()) {
      const date = n => new Date(Date.now() + n * 864e5).toISOString().slice(0, 10);

      await Event.insertMany([
        ['Tech Fest 2026', 'Technical', 'Hackathon, paper presentations and robotics.', 10],
        ['Cultural Day', 'Cultural', 'Dance, music and drama competitions.', 20],
        ['Sports Meet', 'Sports', 'Inter-department athletics and team games.', 30]
      ].map(([title, category, description, n]) => ({
        title,
        category,
        description,
        date: date(n),
        time: '10:00 AM',
        venue: 'Main Auditorium',
        seats: 200,
        deadline: date(n - 1),
        coordinator: 'Event Committee'
      })));

      console.log('Sample events created');
    }

    await cfg();

    app.listen(PORT, () => {
      console.log('Server ready on port ' + PORT);
    });
  } catch (e) {
    console.error('MongoDB connection failed:', e.message);
    process.exit(1);
  }
}

startServer();
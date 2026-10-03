/**
 * SchausHaus Photo Hunt — backend
 *
 * Script Properties required:
 *   SHEET_ID          id of the backing spreadsheet
 *   DRIVE_FOLDER_ID   id of the folder photos are written to
 *   ADMIN_KEY         any random string; ?key=<this> unlocks the judge view
 *   PUBLIC_URL        the link handed out in texts (optional but recommended)
 *
 * A round moves through four phases:
 *   upcoming  prompts sealed on the server, nothing sent to the browser
 *   open      submissions accepted, photos hidden from everyone
 *   voting    photos revealed, one vote per prompt, tallies hidden
 *   final     votes counted, bonus points awarded
 *
 * Run setUp() once after filling in the properties. If you are upgrading an
 * existing season, setUp() is safe to run again — it adds what is missing.
 */

const PROPS = PropertiesService.getScriptProperties();
const TZ = Session.getScriptTimeZone();

const PTS_PER_SUBMISSION = 3;
const PTS_POPULAR_VOTE = 2;
const PTS_JUDGES_CHOICE = 2;
const PTS_FULL_BALLOT = 1;
const DROP_LOWEST_AFTER = 3;
const DEFAULT_VOTING_HOURS = 48;

/**
 * One place that decides whether a key is the judge key. Trimmed on both
 * sides — a trailing space pasted into Script Properties is invisible and
 * would otherwise lock you out of your own game.
 */
function isAdminKey_(key) {
  const real = PROPS.getProperty('ADMIN_KEY');
  if (!real) return false;
  return String(key || '').trim() === String(real).trim();
}

/* ---------------------------------------------------------------- routing */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile('Index');
  t.playerSlug = (p.p || '').toLowerCase();
  t.adminKey = isAdminKey_(p.key) ? String(p.key).trim() : '';
  return t.evaluate()
    .setTitle('SchausHaus Photo Hunt')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ------------------------------------------------------------ sheet utils */

function ss_() {
  return SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
}

function tab_(name) {
  const sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('Missing tab: ' + name + '. Run setUp() first.');
  return sheet;
}

function rows_(name) {
  const values = tab_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values.shift().map(String);
  return values
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

function appendRow_(name, obj) {
  const sheet = tab_(name);
  const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(head.map(function (h) {
    return obj[h] === undefined ? '' : obj[h];
  }));
}

/* -------------------------------------------------------------- first run */

function setUp() {
  const book = ss_();
  const schema = {
    Players: ['slug', 'name', 'active'],
    Rounds: ['roundId', 'title', 'opensAt', 'closesAt', 'votesCloseAt', 'multiplier'],
    Prompts: ['promptId', 'roundId', 'tag', 'text'],
    Submissions: ['subId', 'promptId', 'slug', 'submittedAt', 'fileId', 'thumbId', 'caption'],
    Votes: ['voteId', 'promptId', 'voterSlug', 'choiceSlug', 'castAt'],
    Awards: ['promptId', 'slug', 'kind', 'note']
  };

  Object.keys(schema).forEach(function (name) {
    let sheet = book.getSheetByName(name);
    if (!sheet) sheet = book.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(schema[name]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, schema[name].length).setFontWeight('bold');
      return;
    }
    // Add any columns a previous version did not have.
    // getLastColumn() reports the last column *with data*, so a freshly
    // inserted empty column does not move it. Capture the index first and
    // write to lastCol + 1, or you overwrite the header you meant to keep.
    const head = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    schema[name].forEach(function (col) {
      if (head.indexOf(col) === -1) {
        const lastCol = sheet.getLastColumn();
        sheet.insertColumnAfter(lastCol);
        sheet.getRange(1, lastCol + 1).setValue(col).setFontWeight('bold');
      }
    });
  });

  if (rows_('Players').length === 0) {
    [
      ['chris', 'Chris', true],
      ['gavyn', 'Gavyn', true],
      ['esme', 'Esme', true],
      ['olivia', 'Olivia', true],
      ['bethany', 'Bethany', true],
      ['amanda', 'Amanda', true]
    ].forEach(function (r) { tab_('Players').appendRow(r); });
  }

  const blank = book.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) book.deleteSheet(blank);

  return 'Ready. Add rounds and prompts, then share the link.';
}

/**
 * Four consecutive weekends: Friday 5pm open, Sunday 8pm close,
 * Tuesday 8pm votes close. Final round is double points.
 */
function seedSeason(firstFridayIso) {
  const start = firstFridayIso ? new Date(firstFridayIso) : new Date();
  const titles = ['Base Camp', 'Evidence of Life', 'Off the Trail', 'The Summit'];

  titles.forEach(function (title, i) {
    const opens = new Date(start.getTime());
    opens.setDate(opens.getDate() + i * 7);
    opens.setHours(17, 0, 0, 0);

    const closes = new Date(opens.getTime());
    closes.setDate(closes.getDate() + 2);
    closes.setHours(20, 0, 0, 0);

    const votesClose = new Date(closes.getTime());
    votesClose.setDate(votesClose.getDate() + 2);
    votesClose.setHours(20, 0, 0, 0);

    appendRow_('Rounds', {
      roundId: 'R' + (i + 1),
      title: title,
      opensAt: opens,
      closesAt: closes,
      votesCloseAt: votesClose,
      multiplier: i === titles.length - 1 ? 2 : 1
    });
  });

  const prompts = [
    ['R1', 'FIELD', 'Something within arm\u2019s reach that says more about you than you\u2019d like'],
    ['R1', 'EXPEDITION', 'The furthest point from your bed you reached today'],
    ['R1', 'ODDITY', 'Proof that you are, in fact, an adult'],
    ['R2', 'FIELD', 'The most questionable thing in your fridge'],
    ['R2', 'EXPEDITION', 'A door you have never opened before'],
    ['R2', 'ODDITY', 'Something that should not be where it is'],
    ['R3', 'FIELD', 'A texture worth touching'],
    ['R3', 'EXPEDITION', 'A dog you do not own (ask first)'],
    ['R3', 'ODDITY', 'The saddest plant you can find'],
    ['R4', 'FIELD', 'Something older than Dad'],
    ['R4', 'EXPEDITION', 'The best view you can reach on foot in 20 minutes'],
    ['R4', 'ODDITY', 'A photo that sums up your entire month in one frame']
  ];

  prompts.forEach(function (p, i) {
    appendRow_('Prompts', {
      promptId: p[0] + '-' + (i % 3 + 1),
      roundId: p[0],
      tag: p[1],
      text: p[2]
    });
  });

  return 'Seeded ' + titles.length + ' rounds.';
}

/* ------------------------------------------------------------ round model */

/** Normalises a Rounds row, filling in votesCloseAt if the column is blank. */
function readRound_(r) {
  const closes = new Date(r.closesAt).getTime();
  const fallback = closes + DEFAULT_VOTING_HOURS * 3600 * 1000;

  let votesClose = fallback;
  if (r.votesCloseAt !== '' && r.votesCloseAt !== null && r.votesCloseAt !== undefined) {
    const parsed = new Date(r.votesCloseAt).getTime();
    // A deadline before the round even closed is not a deadline — it is a
    // stray number in the wrong column. Ignore it rather than ending voting
    // before it starts.
    if (!isNaN(parsed) && parsed > closes) votesClose = parsed;
  }

  return {
    roundId: r.roundId,
    title: r.title,
    opensAt: new Date(r.opensAt).getTime(),
    closesAt: closes,
    votesCloseAt: votesClose,
    multiplier: Number(r.multiplier) || 1
  };
}

function phaseOf_(round, nowMs) {
  if (nowMs < round.opensAt) return 'upcoming';
  if (nowMs < round.closesAt) return 'open';
  if (nowMs < round.votesCloseAt) return 'voting';
  return 'final';
}

function allRounds_() {
  return rows_('Rounds').map(readRound_).sort(function (a, b) {
    return a.opensAt - b.opensAt;
  });
}

/**
 * Tallies votes for one round. Ties are not broken — everyone tied takes the
 * bonus, which is the right call for a family game.
 */
function tallyRound_(roundId, promptIds) {
  const inRound = {};
  promptIds.forEach(function (id) { inRound[id] = true; });

  const counts = {};
  rows_('Votes').forEach(function (v) {
    if (!inRound[v.promptId]) return;
    if (!counts[v.promptId]) counts[v.promptId] = {};
    const c = String(v.choiceSlug).toLowerCase();
    counts[v.promptId][c] = (counts[v.promptId][c] || 0) + 1;
  });

  const out = {};
  promptIds.forEach(function (id) {
    const t = counts[id] || {};
    let best = 0;
    Object.keys(t).forEach(function (s) { if (t[s] > best) best = t[s]; });
    out[id] = {
      counts: t,
      winners: best === 0 ? [] : Object.keys(t).filter(function (s) { return t[s] === best; })
    };
  });
  return out;
}

/* ------------------------------------------------------------- read state */

function getState(slug, adminKey) {
  const nowMs = new Date().getTime();
  const isAdmin = isAdminKey_(adminKey);
  const me_ = String(slug || '').toLowerCase();

  const players = rows_('Players').filter(function (p) { return p.active !== false; });
  const rounds = allRounds_();
  const prompts = rows_('Prompts');
  const subs = rows_('Submissions');
  const awards = rows_('Awards');
  const votes = rows_('Votes');

  const promptRound = {};
  prompts.forEach(function (p) { promptRound[p.promptId] = p.roundId; });

  // The round worth showing: live one, else the next up, else the last finished.
  let current = null;
  for (let i = 0; i < rounds.length; i++) {
    const ph = phaseOf_(rounds[i], nowMs);
    if (ph === 'open' || ph === 'voting') { current = rounds[i]; break; }
  }
  if (!current) {
    const upcoming = rounds.filter(function (r) { return r.opensAt > nowMs; });
    const past = rounds.filter(function (r) { return r.votesCloseAt <= nowMs; });
    current = upcoming.length ? upcoming[0] : (past.length ? past[past.length - 1] : null);
  }

  const finalRounds = rounds.filter(function (r) { return phaseOf_(r, nowMs) === 'final'; });

  // Popular-vote bonuses only exist once a round's voting has closed.
  const popularWinners = {};
  finalRounds.forEach(function (r) {
    const ids = prompts.filter(function (p) { return p.roundId === r.roundId; })
      .map(function (p) { return p.promptId; });
    const t = tallyRound_(r.roundId, ids);
    Object.keys(t).forEach(function (pid) { popularWinners[pid] = t[pid].winners; });
  });

  const byPlayer = {};
  players.forEach(function (p) { byPlayer[p.slug] = { rounds: {}, stamps: 0 }; });

  subs.forEach(function (s) {
    const rid = promptRound[s.promptId];
    const who = byPlayer[String(s.slug).toLowerCase()];
    if (!rid || !who) return;
    who.rounds[rid] = (who.rounds[rid] || 0) + PTS_PER_SUBMISSION;
    who.stamps += 1;
  });

  Object.keys(popularWinners).forEach(function (pid) {
    const rid = promptRound[pid];
    popularWinners[pid].forEach(function (s) {
      const who = byPlayer[s];
      if (rid && who) who.rounds[rid] = (who.rounds[rid] || 0) + PTS_POPULAR_VOTE;
    });
  });

  awards.forEach(function (a) {
    const rid = promptRound[a.promptId];
    const who = byPlayer[String(a.slug).toLowerCase()];
    if (!rid || !who) return;
    who.rounds[rid] = (who.rounds[rid] || 0) + PTS_JUDGES_CHOICE;
  });

  // Full-ballot bonus, counted once voting on that round has closed.
  finalRounds.forEach(function (r) {
    const ids = prompts.filter(function (p) { return p.roundId === r.roundId; })
      .map(function (p) { return p.promptId; });
    players.forEach(function (p) {
      const eligible = ids.filter(function (pid) {
        return subs.some(function (s) {
          return s.promptId === pid && String(s.slug).toLowerCase() !== p.slug;
        });
      });
      if (!eligible.length) return;
      const cast = eligible.filter(function (pid) {
        return votes.some(function (v) {
          return v.promptId === pid && String(v.voterSlug).toLowerCase() === p.slug;
        });
      });
      if (cast.length === eligible.length) {
        byPlayer[p.slug].rounds[r.roundId] =
          (byPlayer[p.slug].rounds[r.roundId] || 0) + PTS_FULL_BALLOT;
      }
    });
  });

  const standings = players.map(function (p) {
    const rec = byPlayer[p.slug];
    const perRound = rounds.map(function (r) {
      return {
        roundId: r.roundId,
        points: (rec.rounds[r.roundId] || 0) * r.multiplier,
        counted: true
      };
    });

    if (finalRounds.length >= DROP_LOWEST_AFTER) {
      let worst = null;
      perRound.forEach(function (pr) {
        const done = finalRounds.some(function (f) { return f.roundId === pr.roundId; });
        if (!done) return;
        if (worst === null || pr.points < worst.points) worst = pr;
      });
      if (worst) worst.counted = false;
    }

    return {
      slug: p.slug,
      name: p.name,
      perRound: perRound,
      total: perRound.reduce(function (sum, pr) {
        return sum + (pr.counted ? pr.points : 0);
      }, 0),
      stamps: rec.stamps
    };
  }).sort(function (a, b) { return b.total - a.total || a.name.localeCompare(b.name); });

  const phase = current ? phaseOf_(current, nowMs) : null;

  const mine = {};
  subs.forEach(function (s) {
    if (String(s.slug).toLowerCase() === me_) mine[s.promptId] = true;
  });

  const currentPrompts = !current ? [] : prompts
    .filter(function (p) { return p.roundId === current.roundId; })
    .map(function (p) {
      const sealed = phase === 'upcoming';
      return {
        promptId: p.promptId,
        tag: sealed ? 'SEALED' : p.tag,
        text: sealed ? '' : p.text,
        submitted: !!mine[p.promptId]
      };
    });

  // How much of this round's ballot is still outstanding.
  let ballotLeft = 0;
  if (current && phase === 'voting' && me_) {
    currentPrompts.forEach(function (p) {
      const eligible = subs.some(function (s) {
        return s.promptId === p.promptId && String(s.slug).toLowerCase() !== me_;
      });
      const voted = votes.some(function (v) {
        return v.promptId === p.promptId && String(v.voterSlug).toLowerCase() === me_;
      });
      if (eligible && !voted) ballotLeft += 1;
    });
  }

  const me = players.filter(function (p) { return p.slug === me_; })[0] || null;

  const roundsOut = rounds.map(function (r) {
    return {
      roundId: r.roundId,
      title: r.title,
      opensAt: r.opensAt,
      closesAt: r.closesAt,
      votesCloseAt: r.votesCloseAt,
      multiplier: r.multiplier,
      phase: phaseOf_(r, nowMs)
    };
  });

  return {
    now: nowMs,
    me: me ? { slug: me.slug, name: me.name } : null,
    isAdmin: isAdmin,
    players: players.map(function (p) { return { slug: p.slug, name: p.name }; }),
    rounds: roundsOut,
    current: current,
    phase: phase,
    currentPrompts: currentPrompts,
    ballotLeft: ballotLeft,
    standings: standings
  };
}

/**
 * Gallery for one round. Sealed until submissions close; after that the photos
 * are open to everyone. During voting, each item carries your ballot state.
 */
function getGallery(roundId, slug, adminKey) {
  const nowMs = new Date().getTime();
  const isAdmin = isAdminKey_(adminKey);
  const voter = String(slug || '').toLowerCase();

  const raw = rows_('Rounds').filter(function (r) { return r.roundId === roundId; })[0];
  if (!raw) return { phase: null, items: [] };
  const round = readRound_(raw);
  const phase = phaseOf_(round, nowMs);

  if ((phase === 'upcoming' || phase === 'open') && !isAdmin) {
    return { phase: phase, votesCloseAt: round.votesCloseAt, items: [] };
  }

  const prompts = rows_('Prompts').filter(function (p) { return p.roundId === roundId; });
  const promptText = {};
  prompts.forEach(function (p) { promptText[p.promptId] = p; });
  const promptIds = prompts.map(function (p) { return p.promptId; });

  const names = {};
  rows_('Players').forEach(function (p) { names[p.slug] = p.name; });

  const awarded = {};
  rows_('Awards').forEach(function (a) {
    awarded[a.promptId + '|' + String(a.slug).toLowerCase()] = a.kind;
  });

  // Only ever your own ballot — never anyone else's.
  const myVote = {};
  rows_('Votes').forEach(function (v) {
    if (String(v.voterSlug).toLowerCase() === voter) {
      myVote[v.promptId] = String(v.choiceSlug).toLowerCase();
    }
  });

  const finished = phase === 'final';
  const tally = finished ? tallyRound_(roundId, promptIds) : null;

  const items = rows_('Submissions')
    .filter(function (s) { return !!promptText[s.promptId]; })
    .map(function (s) {
      const owner = String(s.slug).toLowerCase();
      const t = tally && tally[s.promptId];
      return {
        subId: s.subId,
        promptId: s.promptId,
        tag: promptText[s.promptId].tag,
        prompt: promptText[s.promptId].text,
        slug: owner,
        name: names[owner] || owner,
        caption: s.caption || '',
        submittedAt: new Date(s.submittedAt).getTime(),
        thumbUrl: 'https://drive.google.com/thumbnail?id=' + s.thumbId + '&sz=w600',
        fullUrl: 'https://drive.google.com/file/d/' + s.fileId + '/view',
        award: awarded[s.promptId + '|' + owner] || '',
        mine: owner === voter,
        votedFor: myVote[s.promptId] === owner,
        votes: t ? (t.counts[owner] || 0) : null,
        popular: t ? t.winners.indexOf(owner) !== -1 : false
      };
    })
    .sort(function (a, b) {
      return a.promptId.localeCompare(b.promptId) || a.submittedAt - b.submittedAt;
    });

  return { phase: phase, votesCloseAt: round.votesCloseAt, items: items };
}

/* ------------------------------------------------------------------ write */

function submitPhoto(payload) {
  const slug = String(payload.slug || '').toLowerCase();
  const promptId = String(payload.promptId || '');
  if (!slug || !promptId) throw new Error('Missing player or prompt.');

  const player = rows_('Players').filter(function (p) { return p.slug === slug; })[0];
  if (!player) throw new Error('That link does not match a player.');

  const prompt = rows_('Prompts').filter(function (p) { return p.promptId === promptId; })[0];
  if (!prompt) throw new Error('That prompt no longer exists.');

  const round = readRound_(
    rows_('Rounds').filter(function (r) { return r.roundId === prompt.roundId; })[0]);
  const nowMs = new Date().getTime();
  if (nowMs < round.opensAt) throw new Error('This round has not opened yet.');
  if (nowMs >= round.closesAt) throw new Error('This round closed. Nothing more counts.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const existing = rows_('Submissions').filter(function (s) {
      return s.promptId === promptId && String(s.slug).toLowerCase() === slug;
    });
    if (existing.length) throw new Error('You already logged this one. One photo per prompt.');

    const folder = DriveApp.getFolderById(PROPS.getProperty('DRIVE_FOLDER_ID'));
    const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
    const base = prompt.roundId + '_' + prompt.tag + '_' + player.name + '_' + stamp;

    const fullFile = folder.createFile(Utilities.newBlob(
      Utilities.base64Decode(payload.full), 'image/jpeg', base + '.jpg'));
    const thumbFile = folder.createFile(Utilities.newBlob(
      Utilities.base64Decode(payload.thumb), 'image/jpeg', base + '_thumb.jpg'));

    let shareFailed = false;
    [fullFile, thumbFile].forEach(function (f) {
      try {
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (err) {
        // Drive policy blocked link sharing. The photo is safe and the gallery
        // will fall back to serving bytes through the web app, but flag it so
        // this is discoverable rather than silent.
        shareFailed = true;
      }
    });
    if (shareFailed) {
      PROPS.setProperty('SHARING_BLOCKED', 'yes');
    }

    appendRow_('Submissions', {
      subId: Utilities.getUuid().slice(0, 8),
      promptId: promptId,
      slug: slug,
      submittedAt: new Date(),
      fileId: fullFile.getId(),
      thumbId: thumbFile.getId(),
      caption: String(payload.caption || '').slice(0, 200)
    });
  } finally {
    lock.releaseLock();
  }

  return getState(slug, payload.adminKey);
}

/**
 * One vote per prompt. You cannot vote for yourself, and you can change your
 * mind right up until voting closes.
 */
function castVote(promptId, choiceSlug, voterSlug) {
  const voter = String(voterSlug || '').toLowerCase();
  const choice = String(choiceSlug || '').toLowerCase();
  if (!voter || !choice) throw new Error('Missing voter or choice.');
  if (voter === choice) throw new Error('You cannot vote for your own photo.');

  if (!rows_('Players').some(function (p) { return p.slug === voter; })) {
    throw new Error('Tap your name first.');
  }

  const prompt = rows_('Prompts').filter(function (p) { return p.promptId === promptId; })[0];
  if (!prompt) throw new Error('That prompt no longer exists.');

  const round = readRound_(
    rows_('Rounds').filter(function (r) { return r.roundId === prompt.roundId; })[0]);
  const phase = phaseOf_(round, new Date().getTime());
  if (phase === 'upcoming' || phase === 'open') {
    throw new Error('Voting opens when submissions close.');
  }
  if (phase === 'final') throw new Error('Voting is closed for this round.');

  const eligible = rows_('Submissions').some(function (s) {
    return s.promptId === promptId && String(s.slug).toLowerCase() === choice;
  });
  if (!eligible) throw new Error('No photo from that player for this prompt.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = tab_('Votes');
    const values = sheet.getDataRange().getValues();
    const head = values[0].map(String);
    const pi = head.indexOf('promptId');
    const vi = head.indexOf('voterSlug');
    for (let i = values.length - 1; i >= 1; i--) {
      if (values[i][pi] === promptId && String(values[i][vi]).toLowerCase() === voter) {
        sheet.deleteRow(i + 1);
      }
    }
    appendRow_('Votes', {
      voteId: Utilities.getUuid().slice(0, 8),
      promptId: promptId,
      voterSlug: voter,
      choiceSlug: choice,
      castAt: new Date()
    });
  } finally {
    lock.releaseLock();
  }

  return true;
}

/**
 * Checks a judge key. Returns true/false rather than throwing so the sign-in
 * form can show a clean message.
 */
function verifyAdmin(key) {
  if (!PROPS.getProperty('ADMIN_KEY')) {
    throw new Error('No ADMIN_KEY is set in Script Properties yet.');
  }
  return isAdminKey_(key);
}

/**
 * Re-shares every photo already in the sheet and reports what happened.
 * Run this from the editor if people say they cannot see the pictures — the
 * usual cause is a Drive policy that blocked link sharing at upload time.
 */
function repairSharing() {
  const subs = rows_('Submissions');
  let fixed = 0;
  const failures = [];

  subs.forEach(function (s) {
    [s.fileId, s.thumbId].forEach(function (id) {
      if (!id) return;
      try {
        DriveApp.getFileById(id)
          .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fixed += 1;
      } catch (err) {
        failures.push(s.slug + ' / ' + s.promptId + ': ' + err.message);
      }
    });
  });

  const report = 'Re-shared ' + fixed + ' files across ' + subs.length + ' submissions.'
    + (failures.length
        ? '\n\nFAILED (' + failures.length + '):\n' + failures.join('\n')
          + '\n\nLink sharing is blocked on this account. The app will fall back to '
          + 'serving images through the web app itself, which always works.'
        : '\n\nNo failures. Tell everyone to reload.');

  Logger.log(report);
  return report;
}

/**
 * Serves a photo's bytes as a data URL. The fallback for when Drive link
 * sharing is blocked — slower, but it cannot be broken by a sharing policy.
 */
function getPhotoData(subId, big) {
  const sub = rows_('Submissions').filter(function (s) { return s.subId === subId; })[0];
  if (!sub) throw new Error('No such photo.');

  const prompt = rows_('Prompts').filter(function (p) {
    return p.promptId === sub.promptId;
  })[0];
  if (!prompt) throw new Error('No such prompt.');

  // Same seal as the gallery: nothing leaves the server before submissions close.
  const round = readRound_(
    rows_('Rounds').filter(function (r) { return r.roundId === prompt.roundId; })[0]);
  const phase = phaseOf_(round, new Date().getTime());
  if (phase === 'upcoming' || phase === 'open') throw new Error('Round is still sealed.');

  const id = big ? sub.fileId : sub.thumbId;
  const blob = DriveApp.getFileById(id).getBlob();
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
}

/**
 * Pushes a round's voting deadline out. Use it when voting expired before
 * anyone got a chance — extendVoting('R1', 48) gives another two days from now.
 * Pass hours from *now*, not from the original deadline.
 */
function extendVoting(roundId, hours) {
  const sheet = tab_('Rounds');
  const values = sheet.getDataRange().getValues();
  const head = values[0].map(String);
  const idCol = head.indexOf('roundId');
  let col = head.indexOf('votesCloseAt');

  if (col === -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    col = sheet.getLastColumn() - 1;
    sheet.getRange(1, col + 1).setValue('votesCloseAt').setFontWeight('bold');
  }

  const when = new Date(new Date().getTime() + (Number(hours) || 48) * 3600 * 1000);

  for (let i = 1; i < values.length; i++) {
    if (values[i][idCol] === roundId) {
      sheet.getRange(i + 1, col + 1).setValue(when);
      return 'Voting on ' + roundId + ' now runs until '
        + Utilities.formatDate(when, TZ, 'EEEE d MMM, h:mm a') + '.';
    }
  }
  throw new Error('No round called ' + roundId + '.');
}

/**
 * Prints what the server currently believes about every round. Run this first
 * whenever the app is not showing what you expect.
 */
function diagnose() {
  const nowMs = new Date().getTime();
  const lines = ['Server time: ' + Utilities.formatDate(new Date(), TZ, 'EEE d MMM h:mm a'),
                 'Timezone: ' + TZ, ''];

  allRounds_().forEach(function (r) {
    lines.push(r.roundId + '  [' + phaseOf_(r, nowMs).toUpperCase() + ']  ' + r.title);
    lines.push('   opens      ' + Utilities.formatDate(new Date(r.opensAt), TZ, 'EEE h:mm a'));
    lines.push('   closes     ' + Utilities.formatDate(new Date(r.closesAt), TZ, 'EEE h:mm a'));
    lines.push('   votes end  ' + Utilities.formatDate(new Date(r.votesCloseAt), TZ, 'EEE h:mm a'));
  });

  lines.push('', 'Submissions: ' + rows_('Submissions').length,
             'Votes cast:  ' + rows_('Votes').length,
             'Players:     ' + rows_('Players').length);

  const out = lines.join('\n');
  Logger.log(out);
  return out;
}

/**
 * Repairs the Rounds tab after the header-overwrite bug in an earlier setUp().
 * Symptom: the votesCloseAt column holds small numbers (1, 1, 1, 2) instead of
 * dates — those are the multipliers, and the multiplier header was clobbered.
 *
 * Safe to run more than once. It inspects before it writes and reports what it
 * found either way.
 */
function repairRounds() {
  const sheet = tab_('Rounds');
  const values = sheet.getDataRange().getValues();
  const head = values[0].map(String);
  const notes = [];

  const vcCol = head.indexOf('votesCloseAt');
  if (vcCol === -1) {
    notes.push('No votesCloseAt column at all — adding one.');
  }

  // Does the votesCloseAt column actually hold dates?
  let looksLikeMultiplier = false;
  if (vcCol !== -1 && values.length > 1) {
    looksLikeMultiplier = values.slice(1).every(function (row) {
      const v = row[vcCol];
      if (v === '' || v === null) return false;
      return typeof v === 'number' && v >= 0 && v <= 10;
    });
  }

  if (looksLikeMultiplier) {
    if (head.indexOf('multiplier') !== -1 && head.indexOf('multiplier') !== vcCol) {
      notes.push('A separate multiplier column already exists; leaving it alone.');
    } else {
      sheet.getRange(1, vcCol + 1).setValue('multiplier').setFontWeight('bold');
      notes.push('Column ' + String.fromCharCode(65 + vcCol) +
        ' relabelled back to multiplier (its values were 1/2, not dates).');
    }
  } else if (vcCol !== -1) {
    notes.push('votesCloseAt already holds proper values — nothing to undo.');
  }

  // Make sure a real votesCloseAt exists, appended correctly this time.
  const head2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  let target = head2.indexOf('votesCloseAt');
  if (target === -1) {
    const lastCol = sheet.getLastColumn();
    sheet.insertColumnAfter(lastCol);
    target = lastCol;
    sheet.getRange(1, target + 1).setValue('votesCloseAt').setFontWeight('bold');
    notes.push('Added a votesCloseAt column at ' + String.fromCharCode(65 + target) + '.');
  }

  // Fill any blank deadline with 48 hours after that round's close.
  const closeCol = head2.indexOf('closesAt');
  const idCol = head2.indexOf('roundId');
  const rows = sheet.getLastRow();
  let filled = 0;

  for (let r = 2; r <= rows; r++) {
    const id = sheet.getRange(r, idCol + 1).getValue();
    if (!id) continue;
    const cell = sheet.getRange(r, target + 1);
    const existing = cell.getValue();
    const isDate = existing instanceof Date && !isNaN(existing.getTime());
    if (isDate) continue;

    const closes = new Date(sheet.getRange(r, closeCol + 1).getValue());
    if (isNaN(closes.getTime())) continue;
    cell.setValue(new Date(closes.getTime() + DEFAULT_VOTING_HOURS * 3600 * 1000));
    filled += 1;
  }

  if (filled) notes.push('Set a voting deadline on ' + filled + ' round(s): 48h after close.');

  const out = notes.join('\n') + '\n\n' + diagnose();
  Logger.log(out);
  return out;
}

/** Judge's Choice. One per prompt; awarding again replaces the old pick. */
function setAward(promptId, slug, adminKey) {
  if (!isAdminKey_(adminKey)) throw new Error('Judge view only.');
  const sheet = tab_('Awards');
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (values[i][0] === promptId) sheet.deleteRow(i + 1);
  }
  if (slug) {
    appendRow_('Awards', {
      promptId: promptId, slug: String(slug).toLowerCase(), kind: 'judges-choice', note: ''
    });
  }
  return true;
}

/* ------------------------------------------------ text drafts for pasting */

function publicUrl_() {
  const override = PROPS.getProperty('PUBLIC_URL');
  if (override) return override.trim();
  return ScriptApp.getService().getUrl();
}

function when_(ms) {
  return Utilities.formatDate(new Date(ms), TZ, 'EEEE h:mm a');
}

/**
 * The four messages for a round: Friday drop, Saturday nudge, Sunday reveal
 * that opens voting, and Tuesday results.
 */
function draftTexts(roundId, adminKey) {
  if (!isAdminKey_(adminKey)) throw new Error('Judge view only.');

  const state = getState('chris', PROPS.getProperty('ADMIN_KEY'));
  const round = state.rounds.filter(function (r) { return r.roundId === roundId; })[0];
  if (!round) throw new Error('No such round.');

  const prompts = rows_('Prompts').filter(function (p) { return p.roundId === roundId; });
  const promptIds = prompts.map(function (p) { return p.promptId; });
  const link = publicUrl_();

  const drop = [
    'SCHAUSHAUS PHOTO HUNT \u2014 ' + round.roundId + ': ' + String(round.title).toUpperCase(),
    '',
    prompts.map(function (p) { return p.tag + ' \u2014 ' + p.text; }).join('\n'),
    '',
    'Three prompts. Play all three or pick your battles. 3 points each.',
    'Submissions close ' + when_(round.closesAt) + '. Then you all vote.',
    '',
    link
  ].join('\n');

  const submitted = {};
  rows_('Submissions').forEach(function (s) {
    if (promptIds.indexOf(s.promptId) !== -1) submitted[String(s.slug).toLowerCase()] = true;
  });
  const missing = state.players
    .filter(function (p) { return !submitted[p.slug]; })
    .map(function (p) { return p.name; });

  const nudge = (missing.length
    ? 'Submissions close ' + when_(round.closesAt) + '. Still nothing from: '
      + missing.join(', ') + '. The bar is low. A photo of your desk counts.'
    : 'Everyone has logged at least one. Photos unlock ' + when_(round.closesAt) + '.')
    + '\n\n' + link;

  const reveal = [
    round.roundId + ' is closed. Every photo just unlocked \u2014 go look.',
    '',
    'VOTING IS OPEN. Pick a favourite in each of the three prompts. You cannot vote '
      + 'for yourself. Winners take 2 bonus points, and you get 1 point just for '
      + 'filling out all three.',
    '',
    'Voting closes ' + when_(round.votesCloseAt) + '. Nobody sees the tally until then.',
    '',
    link
  ].join('\n');

  const tally = tallyRound_(roundId, promptIds);
  const names = {};
  rows_('Players').forEach(function (p) { names[p.slug] = p.name; });

  const winnerLines = prompts.map(function (p) {
    const w = tally[p.promptId].winners;
    if (!w.length) return p.tag + ' \u2014 no votes cast';
    const n = tally[p.promptId].counts[w[0]];
    return p.tag + ' \u2014 ' + w.map(function (s) { return names[s] || s; }).join(' & ')
      + ' (' + n + (n === 1 ? ' vote)' : ' votes)');
  }).join('\n');

  const board = state.standings.map(function (s, i) {
    return (i + 1) + '. ' + s.name + ' \u2014 ' + s.total;
  }).join('\n');

  const results = [
    round.roundId + ' is settled. You voted, here is what you decided.',
    '',
    'POPULAR VOTE',
    winnerLines,
    '',
    'SEASON STANDINGS',
    board,
    '',
    link
  ].join('\n');

  return { drop: drop, nudge: nudge, reveal: reveal, results: results };
}

/** The season-opening message. Send once, before Round 1. */
function draftIntro() {
  const link = publicUrl_();
  const rounds = allRounds_();
  const first = rounds[0];
  const opensDay = first
    ? Utilities.formatDate(new Date(first.opensAt), TZ, 'EEEE')
    : 'Friday';

  return [
    'SCHAUSHAUS PHOTO HUNT',
    'Season 1 \u2014 ' + rounds.length + ' weekends, one trophy, zero excuses',
    '',
    'Nobody\u2019s home anymore, so I built us a game.',
    '',
    'Every Friday at 5pm I drop three photo prompts. You have until Sunday 8pm to '
      + 'submit. Three points per photo, nine if you do all three.',
    '',
    'Sunday night every photo unlocks at once and you vote on your favourite in each '
      + 'prompt \u2014 no voting for yourself. Winners take 2 bonus points. You get 1 point '
      + 'just for casting all three votes. Voting closes Tuesday 8pm.',
    '',
    'Play all three prompts or pick one. Your worst weekend gets dropped from your '
      + 'total, so a bad Saturday won\u2019t sink you. Final round is worth double, which '
      + 'means nobody\u2019s out of it until the end.',
    '',
    'Nobody sees anybody\u2019s photos until submissions close \u2014 not even me. No '
      + 'stealing the good idea.',
    '',
    'Here\u2019s the link. It\u2019s the same one every time, and it\u2019ll be at the bottom of '
      + 'every text I send:',
    '',
    link,
    '',
    'Tap your name the first time and it\u2019ll remember you. Do this now: open it, then hit '
      + 'Share \u2192 Add to Home Screen (iPhone) or menu \u2192 Add to Home screen (Android). '
      + 'It becomes an icon and you never have to find this text again.',
    '',
    'Round 1 opens ' + opensDay + '. Prompts stay sealed until then.',
    '',
    '\u2014 Commissioner Dad'
  ].join('\n');
}

function emailDrafts(roundId) {
  const t = draftTexts(roundId, PROPS.getProperty('ADMIN_KEY'));
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'SchausHaus drafts \u2014 ' + roundId,
    body: ['--- FRIDAY DROP ---', t.drop,
           '', '--- SATURDAY NUDGE ---', t.nudge,
           '', '--- SUNDAY REVEAL / VOTING OPENS ---', t.reveal,
           '', '--- TUESDAY RESULTS ---', t.results].join('\n')
  });
}

function emailIntro() {
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'SchausHaus \u2014 season opener text',
    body: draftIntro()
  });
}

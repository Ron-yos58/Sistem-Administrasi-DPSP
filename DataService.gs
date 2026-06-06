function getBootstrapData() {
  const user = assertAuthorized_();
  const requests = listRequestsInternal_({ limit: 50 });
  return serializeValue_({
    app: {
      name: APP_CONFIG.APP_NAME,
      version: APP_CONFIG.APP_VERSION
    },
    user: user,
    options: {
      activityTypes: APP_CONFIG.ACTIVITY_TYPES,
      documentTypes: APP_CONFIG.DOCUMENT_TYPES,
      speakerSubtypes: APP_CONFIG.SPEAKER_SUBTYPES,
      speakerStatuses: APP_CONFIG.SPEAKER_STATUSES,
      faculties: APP_CONFIG.FACULTIES
    },
    summary: buildDashboardSummary_(),
    requests: requests.items,
    references: getReferenceDataInternal_()
  });
}

function listRequests(filters) {
  assertAuthorized_();
  return serializeValue_(listRequestsInternal_(filters || {}));
}

function listRequestsInternal_(filters) {
  const sheet = getSheet_('MASTER');
  const lastRow = lastNonEmptyRowInColumn_(sheet, 1);
  if (lastRow < 2) return { items: [], total: 0 };

  const limit = Math.min(
    Math.max(Number(filters.limit || 100), 1),
    APP_CONFIG.MAX_LIST_ROWS
  );
  const query = text_(filters.query).toLowerCase();
  const type = text_(filters.activityType);
  const status = text_(filters.status);
  const rows = sheet.getRange(2, 1, lastRow - 1, MASTER_HEADERS.length).getValues();
  const items = rows.map(masterRowToDto_).filter(function(item) {
    if (!item.id) return false;
    if (!filters.includeArchived && item.status === 'ARCHIVED') return false;
    if (type && item.activityType !== type) return false;
    if (status && item.status !== status) return false;
    if (query) {
      const haystack = [
        item.id, item.activityType, item.activityName, item.partnerName,
        item.documentTypes.join(' '), item.documentNumber
      ].join(' ').toLowerCase();
      if (haystack.indexOf(query) === -1) return false;
    }
    return true;
  }).sort(function(a, b) {
    return String(b.updatedAt || b.createdAt).localeCompare(String(a.updatedAt || a.createdAt));
  });

  return { items: items.slice(0, limit), total: items.length };
}

function getRequestDetail(requestId) {
  assertAuthorized_();
  const id = text_(requestId);
  const master = getSheet_('MASTER');
  const rowNumber = findRowById_(master, id, 1);
  if (!rowNumber) throw new Error('Permohonan tidak ditemukan: ' + id);

  const row = master.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).getValues()[0];
  return serializeValue_({
    request: masterRowToDto_(row),
    documents: getDocumentsByRequest_(id),
    employees: getEmployeesByRequest_(id),
    generatedFiles: getGeneratedFilesByRequest_(id),
    travel: getTravelByRequestInternal_(id),
    emailPreview: buildEmailPreviewInternal_(id, true)
  });
}

function saveRequest(payload) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const clean = normalizeRequestPayload_(payload || {});
  validateRequestPayload_(clean);

  return withScriptLock_(function() {
    const sheet = getSheet_('MASTER');
    let rowNumber = clean.id ? findRowById_(sheet, clean.id, 1) : 0;

    if (!rowNumber && clean.clientToken) {
      rowNumber = findRowByColumnValue_(sheet, clean.clientToken, masterColumn_('Client Token'));
    }

    const isNew = !rowNumber;
    const id = isNew ? generateRequestId_() : text_(sheet.getRange(rowNumber, 1).getValue());
    const existing = isNew
      ? new Array(MASTER_HEADERS.length).fill('')
      : sheet.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).getValues()[0];

    const currentRevision = Number(existing[masterColumn_('Revision') - 1] || 0);
    if (!isNew && clean.revision !== null && clean.revision !== currentRevision) {
      throw new Error(
        'Data sudah diubah pengguna lain. Muat ulang detail sebelum menyimpan.'
      );
    }

    const revision = currentRevision + 1;
    const now = new Date();
    const row = existing.slice();
    setMaster_(row, 'ID Permohonan', id);
    setMaster_(row, 'Tipe Kegiatan', clean.activityType);
    setMaster_(row, 'Jenis Surat', clean.documents.map(function(doc) { return doc.type; }).join('\n'));
    setMaster_(row, 'Sub-Tipe Kegiatan', clean.speakerSubtype);
    setMaster_(row, 'Status Narasumber', clean.speakerStatus);
    setMaster_(row, 'Fakultas Asal Narasumber', clean.faculties.join(', '));
    setMaster_(row, 'Nomor Surat', clean.documents.map(function(doc) {
      return doc.number ? doc.type + ': ' + doc.number : '';
    }).filter(Boolean).join('\n'));
    setMaster_(row, 'Nomor Surat Masuk', clean.incomingNumber);
    setMaster_(row, 'Tanggal Surat Masuk', clean.incomingDate ? parseIsoDate_(clean.incomingDate) : '');
    setMaster_(row, 'Pengirim Surat Masuk', clean.incomingSender);
    setMaster_(row, 'Perihal Surat Masuk', clean.incomingSubject);
    setMaster_(row, 'Nama Kegiatan', clean.activityName);
    setMaster_(row, 'Nama Mitra', clean.partnerName);
    setMaster_(row, 'Alamat Mitra', clean.partnerAddress);
    setMaster_(row, 'Email Mitra', clean.partnerEmail);
    setMaster_(row, 'Tanggal Surat Dibuat', clean.letterDate ? parseIsoDate_(clean.letterDate) : '');
    setMaster_(row, 'Hari Kegiatan', formatDayRange_(clean.startDate, clean.endDate));
    setMaster_(row, 'Tanggal Kegiatan', formatDateRange_(clean.startDate, clean.endDate));
    setMaster_(row, 'Waktu Kegiatan', clean.activityTime);
    setMaster_(row, 'Tempat Kegiatan', clean.activityPlace);
    setMaster_(row, 'Nama Penandatangan', clean.signerName);
    setMaster_(row, 'NIK Penandatangan', clean.signerId);
    setMaster_(row, 'Jabatan Penandatangan', clean.signerRole);
    setMaster_(row, 'Honor', clean.honor);
    setMaster_(row, 'Perjalanan Dinas', clean.travel);
    setMaster_(row, 'Status Permohonan', clean.status);
    setMaster_(row, 'Diubah Oleh', user.email);
    setMaster_(row, 'Diubah Pada', now);
    setMaster_(row, 'Client Token', clean.clientToken);
    setMaster_(row, 'Tanggal Mulai ISO', clean.startDate);
    setMaster_(row, 'Tanggal Selesai ISO', clean.endDate || clean.startDate);
    setMaster_(row, 'Revision', revision);

    if (isNew) {
      setMaster_(row, 'Dibuat Oleh', user.email);
      setMaster_(row, 'Dibuat Pada', now);
    }

    const documents = replaceDocumentsForRequest_(id, clean.documents, revision);
    replaceEmployeesForRequest_(id, clean.employees);
    applyEmployeeAndRoutingColumns_(row, clean.employees, clean);

    if (isNew) {
      rowNumber = appendDataRow_(sheet, row);
    } else {
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    }

    if (clean.travel === 'Ya') syncTravelDataInternal_(id);
    clearAppCache_();
    logAudit_('SAVE_REQUEST', id, true, {
      revision: revision,
      isNew: isNew,
      documents: documents.length,
      employees: clean.employees.length
    });

    return serializeValue_({
      ok: true,
      id: id,
      revision: revision,
      row: rowNumber,
      request: masterRowToDto_(row),
      documents: documents
    });
  });
}

function archiveRequest(requestId, revision) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(requestId);

  return withScriptLock_(function() {
    const sheet = getSheet_('MASTER');
    const rowNumber = findRowById_(sheet, id, 1);
    if (!rowNumber) throw new Error('Permohonan tidak ditemukan: ' + id);
    const row = sheet.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).getValues()[0];
    const currentRevision = Number(row[masterColumn_('Revision') - 1] || 0);
    if (Number(revision) !== currentRevision) {
      throw new Error('Revision tidak cocok. Muat ulang data.');
    }

    setMaster_(row, 'Status Permohonan', 'ARCHIVED');
    setMaster_(row, 'Diubah Oleh', user.email);
    setMaster_(row, 'Diubah Pada', new Date());
    setMaster_(row, 'Revision', currentRevision + 1);
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    clearAppCache_();
    logAudit_('ARCHIVE_REQUEST', id, true, {});
    return { ok: true };
  });
}

function getReferenceData() {
  assertAuthorized_();
  return serializeValue_(getReferenceDataInternal_());
}

function getReferenceDataInternal_() {
  const ccSheet = getSheet_('CC');
  const lastRow = lastNonEmptyRowInColumn_(ccSheet, 1);
  const cc = lastRow < 2 ? [] :
    ccSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
      .filter(function(row) { return row[1] && row[2]; })
      .map(function(row) {
        return { unit: row[0], role: row[1], email: row[2] };
      });
  return { cc: cc };
}

function buildDashboardSummary_() {
  const requests = listRequestsInternal_({ limit: APP_CONFIG.MAX_LIST_ROWS }).items;
  const documents = readDataRows_(getSheet_('DOCUMENTS'), DOCUMENT_HEADERS.length)
    .map(documentRowToDto_);
  const now = new Date();
  const next30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

  return {
    total: requests.length,
    drafts: requests.filter(function(item) { return item.status === 'DRAFT'; }).length,
    ready: requests.filter(function(item) { return item.status === 'READY'; }).length,
    documentsPending: documents.filter(function(item) {
      return item.status === 'PENDING' || item.status === 'ERROR';
    }).length,
    emailsPending: documents.filter(function(item) {
      return item.status === 'GENERATED' && item.emailStatus !== 'DRAFTED';
    }).length,
    upcoming: requests.filter(function(item) {
      if (!item.startDate) return false;
      const date = parseIsoDate_(item.startDate);
      return date >= now && date <= next30;
    }).length
  };
}

function masterRowToDto_(row) {
  return {
    id: masterValue_(row, 'ID Permohonan'),
    activityType: masterValue_(row, 'Tipe Kegiatan'),
    documentTypes: String(masterValue_(row, 'Jenis Surat') || '').split(/\r?\n/).filter(Boolean),
    speakerSubtype: masterValue_(row, 'Sub-Tipe Kegiatan'),
    speakerStatus: masterValue_(row, 'Status Narasumber'),
    faculties: String(masterValue_(row, 'Fakultas Asal Narasumber') || '').split(/\s*,\s*/).filter(Boolean),
    documentNumber: masterValue_(row, 'Nomor Surat'),
    incomingNumber: masterValue_(row, 'Nomor Surat Masuk'),
    incomingDate: serializeValue_(masterValue_(row, 'Tanggal Surat Masuk')),
    incomingSender: masterValue_(row, 'Pengirim Surat Masuk'),
    incomingSubject: masterValue_(row, 'Perihal Surat Masuk'),
    activityName: masterValue_(row, 'Nama Kegiatan'),
    partnerName: masterValue_(row, 'Nama Mitra'),
    partnerAddress: masterValue_(row, 'Alamat Mitra'),
    partnerEmail: masterValue_(row, 'Email Mitra'),
    letterDate: serializeValue_(masterValue_(row, 'Tanggal Surat Dibuat')),
    day: masterValue_(row, 'Hari Kegiatan'),
    dateDisplay: masterValue_(row, 'Tanggal Kegiatan'),
    startDate: masterValue_(row, 'Tanggal Mulai ISO'),
    endDate: masterValue_(row, 'Tanggal Selesai ISO'),
    activityTime: masterValue_(row, 'Waktu Kegiatan'),
    activityPlace: masterValue_(row, 'Tempat Kegiatan'),
    signerName: masterValue_(row, 'Nama Penandatangan'),
    signerId: masterValue_(row, 'NIK Penandatangan'),
    signerRole: masterValue_(row, 'Jabatan Penandatangan'),
    honor: masterValue_(row, 'Honor') || 'Tidak',
    travel: masterValue_(row, 'Perjalanan Dinas') || 'Tidak',
    emailTo: masterValue_(row, 'Email To'),
    emailCc: masterValue_(row, 'Email CC'),
    emailNote: masterValue_(row, 'Keterangan Email'),
    emailStatus: masterValue_(row, 'Email Status'),
    status: masterValue_(row, 'Status Permohonan') || 'DRAFT',
    createdBy: masterValue_(row, 'Dibuat Oleh'),
    createdAt: serializeValue_(masterValue_(row, 'Dibuat Pada')),
    updatedBy: masterValue_(row, 'Diubah Oleh'),
    updatedAt: serializeValue_(masterValue_(row, 'Diubah Pada')),
    clientToken: masterValue_(row, 'Client Token'),
    revision: Number(masterValue_(row, 'Revision') || 0)
  };
}

function getDocumentsByRequest_(requestId) {
  return readDataRows_(getSheet_('DOCUMENTS'), DOCUMENT_HEADERS.length)
    .filter(function(row) { return text_(row[1]) === requestId; })
    .map(documentRowToDto_);
}

function documentRowToDto_(row) {
  return {
    id: row[0],
    requestId: row[1],
    type: row[2],
    speakerSubtype: row[3],
    speakerStatus: row[4],
    number: row[5],
    templateKey: row[6],
    status: row[7] || 'PENDING',
    docId: row[8],
    docUrl: row[9],
    pdfId: row[10],
    pdfUrl: row[11],
    emailDraftId: row[12],
    emailStatus: row[13],
    revision: Number(row[14] || 0),
    createdAt: serializeValue_(row[15]),
    updatedAt: serializeValue_(row[16])
  };
}

function getEmployeesByRequest_(requestId) {
  return readDataRows_(getSheet_('EMPLOYEES'), EMPLOYEE_HEADERS.length)
    .filter(function(row) { return text_(row[0]) === requestId; })
    .map(function(row) {
      return {
        name: row[1],
        identifier: row[2],
        role: row[3],
        unit: row[4],
        email: row[5],
        rank: row[6],
        category: row[7],
        participantKey: row[8]
      };
    });
}

function replaceDocumentsForRequest_(requestId, documents, revision) {
  const sheet = getSheet_('DOCUMENTS');
  const allRows = readDataRows_(sheet, DOCUMENT_HEADERS.length);
  const existing = allRows.filter(function(row) { return text_(row[1]) === requestId; });
  const untouched = allRows.filter(function(row) { return text_(row[1]) !== requestId; });
  const now = new Date();

  const output = documents.map(function(document) {
    let old = null;
    if (document.id) old = existing.find(function(row) { return text_(row[0]) === document.id; });
    if (!old) old = existing.find(function(row) { return text_(row[2]) === document.type; });

    const templateKey = resolveTemplateKey_({
      activityType: document.activityType,
      type: document.type,
      speakerSubtype: document.speakerSubtype,
      speakerStatus: document.speakerStatus
    });
    const row = old ? old.slice() : new Array(DOCUMENT_HEADERS.length).fill('');
    row[0] = old ? old[0] : 'DOC-' + Utilities.getUuid().slice(0, 12).toUpperCase();
    row[1] = requestId;
    row[2] = document.type;
    row[3] = document.speakerSubtype;
    row[4] = document.speakerStatus;
    row[5] = document.number;
    row[6] = templateKey;
    row[7] = old && old[2] === document.type && old[5] === document.number
      ? (old[7] || 'PENDING')
      : 'PENDING';
    row[14] = revision;
    row[15] = old ? old[15] : now;
    row[16] = now;
    return row;
  });

  rewriteDataRows_(sheet, untouched.concat(output), DOCUMENT_HEADERS.length);
  return output.map(documentRowToDto_);
}

function replaceEmployeesForRequest_(requestId, employees) {
  const sheet = getSheet_('EMPLOYEES');
  const untouched = readDataRows_(sheet, EMPLOYEE_HEADERS.length)
    .filter(function(row) { return text_(row[0]) !== requestId; });
  const additions = employees.map(function(employee) {
    return [
      requestId,
      employee.name,
      employee.identifier,
      employee.role,
      employee.unit,
      employee.email,
      employee.rank,
      employee.category,
      normalizeParticipantKeyForRequest_(requestId, employee.participantKey)
    ];
  });
  rewriteDataRows_(sheet, untouched.concat(additions), EMPLOYEE_HEADERS.length);
}

function normalizeRequestPayload_(payload) {
  const activityType = text_(payload.activityType);
  const speakerSubtype = text_(payload.speakerSubtype);
  const speakerStatus = text_(payload.speakerStatus);
  const documents = (payload.documents || []).map(function(document) {
    return {
      id: text_(document.id),
      activityType: activityType,
      type: text_(document.type),
      speakerSubtype: speakerSubtype,
      speakerStatus: speakerStatus,
      number: text_(document.number, 150)
    };
  });
  const employees = (payload.employees || []).map(function(employee) {
    return {
      name: text_(employee.name, 200),
      identifier: text_(employee.identifier, 100),
      role: text_(employee.role, 200),
      unit: text_(employee.unit, 200),
      email: emailList_(employee.email).join(', '),
      rank: text_(employee.rank, 200),
      category: text_(employee.category, 200),
      participantKey: text_(employee.participantKey) || buildParticipantKey_(payload.id, employee)
    };
  }).filter(function(employee) {
    return employee.name || employee.identifier || employee.email;
  });

  return {
    id: text_(payload.id),
    revision: payload.revision === '' || payload.revision == null ? null : Number(payload.revision),
    clientToken: text_(payload.clientToken) || Utilities.getUuid(),
    status: ['DRAFT', 'READY'].indexOf(text_(payload.status).toUpperCase()) !== -1
      ? text_(payload.status).toUpperCase()
      : 'DRAFT',
    activityType: activityType,
    speakerSubtype: speakerSubtype,
    speakerStatus: speakerStatus,
    faculties: uniqueTextList_(payload.faculties || []),
    incomingNumber: text_(payload.incomingNumber, 150),
    incomingDate: text_(payload.incomingDate),
    incomingSender: text_(payload.incomingSender, 250),
    incomingSubject: text_(payload.incomingSubject, 500),
    activityName: text_(payload.activityName, 300),
    partnerName: text_(payload.partnerName, 300),
    partnerAddress: text_(payload.partnerAddress, 1000),
    partnerEmail: emailList_(payload.partnerEmail).join(', '),
    letterDate: text_(payload.letterDate),
    startDate: text_(payload.startDate),
    endDate: text_(payload.endDate) || text_(payload.startDate),
    activityTime: text_(payload.activityTime, 100),
    activityPlace: text_(payload.activityPlace, 500),
    signerName: text_(payload.signerName, 250),
    signerId: text_(payload.signerId, 100),
    signerRole: text_(payload.signerRole, 250),
    honor: text_(payload.honor) === 'Ya' ? 'Ya' : 'Tidak',
    travel: text_(payload.travel) === 'Ya' ? 'Ya' : 'Tidak',
    documents: documents,
    employees: employees
  };
}

function validateRequestPayload_(payload) {
  const errors = [];
  const isReady = payload.status === 'READY';
  if (APP_CONFIG.ACTIVITY_TYPES.indexOf(payload.activityType) === -1) {
    errors.push('Tipe kegiatan tidak valid.');
  }
  if (isReady && !payload.activityName) errors.push('Nama kegiatan wajib diisi.');
  if (isReady && !payload.partnerName) errors.push('Nama mitra wajib diisi.');
  if (isReady && !payload.startDate) errors.push('Tanggal mulai wajib diisi.');
  if (isReady && !payload.activityPlace) errors.push('Tempat kegiatan wajib diisi.');
  if (isReady && !payload.documents.length) errors.push('Pilih minimal satu dokumen.');

  if (payload.startDate) {
    parseIsoDate_(payload.startDate);
    parseIsoDate_(payload.endDate || payload.startDate);
    formatDateRange_(payload.startDate, payload.endDate);
  }
  if (payload.incomingDate) parseIsoDate_(payload.incomingDate);
  if (payload.letterDate) parseIsoDate_(payload.letterDate);

  const allowed = allowedDocumentsFor_(payload.activityType, payload.speakerSubtype);
  payload.documents.forEach(function(document) {
    if (allowed.indexOf(document.type) === -1) {
      errors.push('Jenis dokumen tidak sesuai kegiatan: ' + document.type);
    }
    if (isReady && !document.number) {
      errors.push('Nomor surat wajib untuk status siap: ' + document.type);
    }
  });

  const needsPeople = payload.documents.some(function(document) {
    return document.type === 'Surat Tugas';
  }) || (
    payload.documents.some(function(document) {
      return document.type === 'Surat Permohonan Narasumber kepada Dekan';
    }) && payload.speakerStatus === 'Tidak Dicarikan'
  );
  if (isReady && needsPeople && !payload.employees.length) {
    errors.push('Minimal satu pegawai/narasumber wajib diisi.');
  }

  if (isReady && payload.activityType === 'Penugasan Narasumber') {
    if (APP_CONFIG.SPEAKER_SUBTYPES.indexOf(payload.speakerSubtype) === -1) {
      errors.push('Sub-tipe narasumber wajib diisi.');
    }
    const hasRequestLetter = payload.documents.some(function(document) {
      return document.type === 'Surat Permohonan Narasumber kepada Dekan';
    });
    if (hasRequestLetter) {
      if (APP_CONFIG.SPEAKER_STATUSES.indexOf(payload.speakerStatus) === -1) {
        errors.push('Status pencarian narasumber wajib diisi.');
      }
      if (!payload.faculties.length) errors.push('Fakultas tujuan wajib dipilih.');
    }
  }

  if (isReady && payload.travel === 'Ya') {
    payload.employees.forEach(function(employee, index) {
      if (!employee.rank || !employee.category) {
        errors.push('Pangkat dan kategori perjadin wajib untuk orang ke-' + (index + 1) + '.');
      }
    });
  }
  if (errors.length) throw new Error(errors.join('\n'));
}

function allowedDocumentsFor_(activityType, speakerSubtype) {
  if (activityType === 'Edu Fair') return ['Surat Tugas'];
  if (activityType === 'Campus Visit') {
    return [
      'Surat Tugas',
      'Surat Balasan Campus Visit',
      'Surat Rekomendasi Campus Visit - SU',
      'Surat izin pimpinan - Campus Visit'
    ];
  }
  if (activityType === 'Penugasan Narasumber') {
    return speakerSubtype === 'Promosi'
      ? ['Surat Tugas']
      : ['Surat Tugas', 'Surat Permohonan Narasumber kepada Dekan'];
  }
  return [];
}

function applyEmployeeAndRoutingColumns_(row, employees, request) {
  const count = employees.length;
  setMaster_(row, 'Nomor Urut Pegawai', Array.from({ length: count }, function(_, i) {
    return i + 1;
  }).join('\n'));
  setMaster_(row, 'Nama Pegawai', employees.map(function(item) { return item.name; }).join('\n'));
  setMaster_(row, 'NIP/NPM Pegawai', employees.map(function(item) { return item.identifier; }).join('\n'));
  setMaster_(row, 'Jabatan Pegawai', employees.map(function(item) { return item.role; }).join('\n'));
  setMaster_(row, 'Prodi/Unit Pegawai', employees.map(function(item) { return item.unit; }).join('\n'));
  setMaster_(row, 'Email Pegawai', employees.map(function(item) { return item.email; }).filter(Boolean).join('\n'));

  const routing = computeEmailRouting_(request, employees);
  setMaster_(row, 'Email To', routing.to.join('\n'));
  setMaster_(row, 'Jabatan Email To', routing.toRoles.join('\n'));
  setMaster_(row, 'Email CC', routing.cc.join('\n'));
  setMaster_(row, 'Jabatan Email CC', routing.ccRoles.join('\n'));
  setMaster_(row, 'Keterangan Email', routing.notes.join(' | '));
}

function computeEmailRouting_(request, employees, documentType) {
  const references = getReferenceDataInternal_().cc;
  const to = [];
  const toRoles = [];
  const cc = [];
  const ccRoles = [];
  const notes = [];
  const types = documentType
    ? [documentType]
    : request.documents.map(function(document) { return document.type; });

  function addByRole(targetEmails, targetRoles, role, unit) {
    const match = references.find(function(item) {
      return item.role === role && (!unit || item.unit === unit);
    });
    if (match) {
      targetEmails.push(match.email);
      targetRoles.push(match.role);
    } else {
      notes.push('Master CC tidak menemukan: ' + role);
    }
  }

  function addEmployees() {
    employees.forEach(function(employee) {
      emailList_(employee.email).forEach(function(email) { to.push(email); });
      if (employee.name) toRoles.push(employee.name);
    });
  }

  types.forEach(function(type) {
    if (type === 'Surat Balasan Campus Visit') {
      emailList_(request.partnerEmail).forEach(function(email) { to.push(email); });
      toRoles.push(request.partnerName);
      addByRole(cc, ccRoles, 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', 'Rektorat');
    } else if (type === 'Surat Rekomendasi Campus Visit - SU') {
      addByRole(to, toRoles, 'Sekretaris Universitas', 'Rektorat');
      addByRole(cc, ccRoles, 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', 'Rektorat');
    } else if (type === 'Surat izin pimpinan - Campus Visit') {
      [
        ['Sekretaris Universitas', 'Rektorat'],
        ['Dekan Fakultas Ekonomi', 'Fakultas Ekonomi'],
        ['Dekan Fakultas Hukum', 'Fakultas Hukum'],
        ['Dekan Fakultas Ilmu Sosial dan Ilmu Politik', 'Fakultas Ilmu Sosial dan Ilmu Politik'],
        ['Dekan Fakultas Teknik', 'Fakultas Teknik'],
        ['Dekan Fakultas Teknologi Rekayasa', 'Fakultas Teknologi Rekayasa'],
        ['Dekan Fakultas Sains', 'Fakultas Sains'],
        ['Dekan Fakultas Vokasi', 'Fakultas Vokasi'],
        ['Direktur Kemahasiswaan', 'Direktorat Kemahasiswaan'],
        ['Direktur Manajemen Aset, Keuangan, dan Sarana Prasarana', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
        ['Kepala Perpustakaan', 'Unit Perpustakaan']
      ].forEach(function(pair) { addByRole(to, toRoles, pair[0], pair[1]); });

      [
        ['Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', 'Rektorat'],
        ['Manajer Aset dan Sarana Prasarana', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
        ['Manajer Kemahasiswaan', 'Direktorat Kemahasiswaan'],
        ['Koordinator Kebersihan, Keamanan dan Ketertiban', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
        ['Koordinator Kelas dan Fasilitas Pendukung', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
        ['Koordinator Administrasi Fakultas Ekonomi', 'Fakultas Ekonomi'],
        ['Koordinator Administrasi Fakultas Hukum', 'Fakultas Hukum'],
        ['Koordinator Administrasi Fakultas Ilmu Sosial dan Ilmu Politik', 'Fakultas Ilmu Sosial dan Ilmu Politik'],
        ['Koordinator Administrasi Fakultas Sains', 'Fakultas Sains'],
        ['Koordinator Administrasi Fakultas Teknik', 'Fakultas Teknik'],
        ['Koordinator Administrasi Fakultas Teknologi Rekayasa', 'Fakultas Teknologi Rekayasa'],
        ['Koordinator Administrasi Fakultas Vokasi', 'Fakultas Vokasi']
      ].forEach(function(pair) { addByRole(cc, ccRoles, pair[0], pair[1]); });
    } else if (type === 'Surat Permohonan Narasumber kepada Dekan') {
      request.faculties.forEach(function(faculty) {
        addByRole(to, toRoles, 'Dekan ' + faculty, faculty);
        addByRole(cc, ccRoles, 'Koordinator Administrasi ' + faculty, faculty);
      });
      if (request.speakerStatus === 'Tidak Dicarikan') addEmployees();
    } else if (type === 'Surat Tugas') {
      addEmployees();
    }
  });

  return {
    to: emailList_(to),
    toRoles: uniqueTextList_(toRoles),
    cc: emailList_(cc),
    ccRoles: uniqueTextList_(ccRoles),
    notes: uniqueTextList_(notes)
  };
}

function resolveTemplateKey_(document) {
  if (document.activityType === 'Edu Fair' && document.type === 'Surat Tugas') {
    return 'EDU_FAIR_TASK';
  }
  if (document.activityType === 'Campus Visit') {
    if (document.type === 'Surat Tugas') return 'CAMPUS_VISIT_TASK';
    if (document.type === 'Surat izin pimpinan - Campus Visit') return 'CAMPUS_VISIT_PERMISSION';
    if (document.type === 'Surat Rekomendasi Campus Visit - SU') return 'CAMPUS_VISIT_RECOMMENDATION';
    if (document.type === 'Surat Balasan Campus Visit') return 'CAMPUS_VISIT_REPLY';
  }
  if (document.activityType === 'Penugasan Narasumber') {
    if (document.type === 'Surat Tugas' && document.speakerSubtype === 'Promosi') {
      return 'SPEAKER_PROMOTION_TASK';
    }
    if (document.type === 'Surat Tugas') return 'SPEAKER_WORKSHOP_TASK';
    if (document.type === 'Surat Permohonan Narasumber kepada Dekan') {
      return document.speakerStatus === 'Dicarikan'
        ? 'SPEAKER_REQUEST_SEARCH'
        : 'SPEAKER_REQUEST_KNOWN';
    }
  }
  throw new Error('Template tidak tersedia untuk kombinasi dokumen.');
}

function masterColumn_(header) {
  const index = MASTER_HEADERS.indexOf(header);
  if (index === -1) throw new Error('Header Master tidak dikenal: ' + header);
  return index + 1;
}

function masterValue_(row, header) {
  return row[masterColumn_(header) - 1];
}

function setMaster_(row, header, value) {
  row[masterColumn_(header) - 1] = value;
}

function readDataRows_(sheet, width) {
  const lastRow = lastNonEmptyRowInColumn_(sheet, 1);
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, width).getValues()
    .filter(function(row) {
      return row.some(function(value) { return value !== ''; });
    });
}

function rewriteDataRows_(sheet, rows, width) {
  const oldRows = Math.max(lastNonEmptyRowInColumn_(sheet, 1) - 1, 0);
  if (oldRows) sheet.getRange(2, 1, oldRows, width).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows);
}

function findRowByColumnValue_(sheet, value, column) {
  const lastRow = lastNonEmptyRowInColumn_(sheet, 1);
  if (!value || lastRow < 2) return 0;
  const values = sheet.getRange(2, column, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (text_(values[i][0]) === text_(value)) return i + 2;
  }
  return 0;
}

function assertCanWrite_(user) {
  if (user.role === 'VIEWER') throw new Error('Role Viewer tidak dapat mengubah data.');
}

function buildParticipantKey_(requestId, employee) {
  const base = text_(employee.identifier) || text_(employee.email) || text_(employee.name);
  const normalized = base.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  return (text_(requestId) || 'NEW') + '|' + (normalized || Utilities.getUuid());
}

function normalizeParticipantKeyForRequest_(requestId, participantKey) {
  const key = text_(participantKey);
  if (!key) return requestId + '|' + Utilities.getUuid();
  if (key.indexOf('NEW|') === 0) return requestId + key.slice(3);
  return key;
}

  function getBootstrapData() {
    const user = assertAuthorized_();
    let requests = getJsonCache_('bootstrap_requests');
    if (!requests) {
      requests = listRequestsInternal_({ limit: 50 });
      putJsonCache_('bootstrap_requests', requests);
    }
    let summary = getJsonCache_('bootstrap_summary');
    if (!summary) {
      summary = buildDashboardSummary_();
      putJsonCache_('bootstrap_summary', summary);
    }
    return serializeValue_({
      app: {
        name: APP_CONFIG.APP_NAME,
        version: APP_CONFIG.APP_VERSION
      },
      user: user,
      options: {
        activityTypes: APP_CONFIG.ACTIVITY_TYPES,
        documentTypes: APP_CONFIG.DOCUMENT_TYPES,
        officialLetterMapping: OFFICIAL_LETTER_MAPPING,
        faculties: APP_CONFIG.FACULTIES,
        signers: getSignatureConfigsInternal_(),
        employeeCatalog: getEmployeeCatalogInternal_()
      },
      summary: summary,
      requests: requests.items,
      references: getReferenceDataInternal_()
    });
  }

  function refreshBootstrapData() {
    assertAuthorized_();
    clearAppCache_();
    return getBootstrapData();
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
    const documentsByRequest = buildDocumentsByRequestMap_();
    const items = rows.map(function(row) {
      return enrichRequestWithDocuments_(masterRowToDto_(row), documentsByRequest);
    }).filter(function(item) {
      if (!item.id) return false;
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
    const cacheKey = 'req_detail_pub_' + id;
    const cached = getJsonCache_(cacheKey);
    if (cached) {
      return serializeValue_(cached);
    }

    const master = getSheet_('MASTER');
    const rowNumber = findRowById_(master, id, 1);
    if (!rowNumber) throw new Error('Permohonan tidak ditemukan: ' + id);

    const row = master.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).getValues()[0];
    const request = masterRowToDto_(row);
    const employees = getEmployeesByRequest_(id);
    const travel = getTravelByRequestInternal_(id);
    const generatedFiles = getGeneratedFilesByRequest_(id);
    const documents = getDocumentsByRequest_(id).map(function(doc) {
      const routingRequest = {
        activityType: request.activityType,
        documents: [doc],
        partnerEmail: request.partnerEmail,
        partnerName: request.partnerName,
        faculties: request.faculties,
        speakerStatus: request.speakerStatus,
        manualTo: [],
        manualCc: Array.isArray(request.manualCc) ? request.manualCc : []
      };
      doc.routing = computeEmailRouting_(routingRequest, employees, doc.type);
      const templateConfig = getTemplateConfigByKey_(doc.templateKey);
      doc.letterCode = templateConfig ? templateConfig.code : 'I';
      doc.letterType = templateConfig ? templateConfig.name : doc.type;
      return doc;
    });
    const financeArtifacts = getFinanceArtifactUrls_(id, generatedFiles);
    const result = {
      request: Object.assign(
        enrichRequestWithDocuments_(request, null, documents),
        financeArtifacts
      ),
      documents: documents,
      schedules: getSchedulesByRequest_(id, row),
      employees: employees,
      generatedFiles: generatedFiles,
      financeArtifacts: financeArtifacts,
      financeReadiness: getFinanceReadiness_(id, { request: request, employees: employees, travel: travel }, generatedFiles),
      travel: travel
    };

    putJsonCache_(cacheKey, result);
    trackCachedRequestId_(id);

    return serializeValue_(result);
  }

  function buildDocumentsByRequestMap_() {
    const rows = readDataRows_(getSheet_('DOCUMENTS'), DOCUMENT_HEADERS.length);
    return rows.reduce(function(map, row) {
      const requestId = text_(row[1]);
      if (!requestId) return map;
      if (!map[requestId]) map[requestId] = [];
      map[requestId].push(documentRowToDto_(row));
      return map;
    }, {});
  }

  function enrichRequestWithDocuments_(request, documentsByRequest, preloadedDocuments) {
    const item = Object.assign({}, request);
    const docs = preloadedDocuments || (documentsByRequest && documentsByRequest[item.id]) || [];
    item.documentProgress = summarizeDocumentWorkflow_(docs, item.status);
    if (docs.length) {
      item.documentTypes = uniqueTextList_(docs.map(function(doc) { return doc.type; }));
      item.documentNumber = docs.map(function(doc) {
        return doc.number ? doc.type + ': ' + doc.number : '';
      }).filter(Boolean).join('\n');
    }
    return item;
  }

  function summarizeDocumentWorkflow_(documents, requestStatus) {
    const docs = documents || [];
    const summary = {
      total: docs.length,
      notCreated: 0,
      generated: 0,
      failed: 0,
      drafted: 0,
      status: 'NOT_CREATED'
    };
    docs.forEach(function(document) {
      if (document.status === 'ERROR') summary.failed++;
      else if (document.status === 'GENERATED') summary.generated++;
      else summary.notCreated++;
      if (document.emailStatus === 'DRAFTED') summary.drafted++;
    });

    if (requestStatus === 'ARCHIVED') summary.status = 'COMPLETE';
    else if (summary.failed) summary.status = 'ERROR';
    else if (!summary.total || summary.notCreated === summary.total) summary.status = 'NOT_CREATED';
    else if (summary.drafted === summary.total) summary.status = 'DRAFTED';
    else if (summary.generated === summary.total) summary.status = 'GENERATED';
    else summary.status = 'IN_PROGRESS';
    return summary;
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
      if (!isNew && text_(masterValue_(existing, 'Status Permohonan')) === 'ARCHIVED') {
        throw new Error('Permohonan selesai bersifat hanya-baca dan tidak dapat diubah.');
      }

      const currentRevision = Number(existing[masterColumn_('Revision') - 1] || 0);
      if (!isNew && clean.revision !== null && clean.revision !== currentRevision) {
        throw new Error(
          'Data sudah diubah pengguna lain. Muat ulang detail sebelum menyimpan.'
        );
      }
      if (!isNew && !clean.documents.length) {
        clean.documents = getDocumentsByRequest_(id).map(function(document) {
          return {
            id: document.id,
            activityType: clean.activityType,
            type: document.type,
            speakerSubtype: document.speakerSubtype,
            speakerStatus: document.speakerStatus,
            number: document.number
          };
        });
      }
      const previousClean = isNew ? null : normalizeRequestPayload_(Object.assign(
        {},
        masterRowToDto_(existing),
        {
          documents: getDocumentsByRequest_(id),
          schedules: getSchedulesByRequest_(id, existing),
          employees: getEmployeesByRequest_(id)
        }
      ));
      const generationChanged = isNew ||
        requestGenerationFingerprint_(previousClean) !== requestGenerationFingerprint_(clean);

      const revision = currentRevision + 1;
      const now = new Date();
      const row = existing.slice();
      const scheduleSummary = buildScheduleSummary_(clean.schedules);
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
      setMaster_(row, 'Hari Kegiatan', scheduleSummary.dayDisplay);
      setMaster_(row, 'Tanggal Kegiatan', scheduleSummary.dateDisplay);
      setMaster_(row, 'Waktu Kegiatan', scheduleSummary.timeDisplay);
      setMaster_(row, 'Tempat Kegiatan', scheduleSummary.placeDisplay);
      setMaster_(row, 'Nama Penandatangan', clean.signerName);
      setMaster_(row, 'NIK Penandatangan', clean.signerId);
      setMaster_(row, 'Jabatan Penandatangan', clean.signerRole);
      setMaster_(row, 'Honor', clean.honor);
      setMaster_(row, 'Perjalanan Dinas', clean.travel);
      setMaster_(row, 'Status Permohonan', clean.status);
      setMaster_(row, 'Diubah Oleh', user.email);
      setMaster_(row, 'Diubah Pada', now);
      setMaster_(row, 'Client Token', clean.clientToken);
      setMaster_(row, 'Tanggal Mulai ISO', scheduleSummary.startDate);
      setMaster_(row, 'Tanggal Selesai ISO', scheduleSummary.endDate);
      setMaster_(row, 'Revision', revision);

      if (isNew) {
        setMaster_(row, 'Dibuat Oleh', user.email);
        setMaster_(row, 'Dibuat Pada', now);
      }

      if (generationChanged) clearGeneratedStateColumns_(row);
      if (isNew) {
        rowNumber = appendDataRow_(sheet, row);
      } else {
        sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
        if (generationChanged) markGeneratedFilesSuperseded_(id);
      }

      const documents = clean.documents.length
        ? replaceDocumentsForRequest_(id, clean.documents, revision, !generationChanged)
        : getDocumentsByRequest_(id);
      replaceSchedulesForRequest_(id, clean.schedules);
      replaceEmployeesForRequest_(id, clean.employees);
      syncDocumentRecipientsForRequest_(id, clean, clean.employees);
      applyEmployeeAndRoutingColumns_(row, clean.employees, clean);
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

      if (clean.status === 'READY') syncTravelDataInternal_(id);
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
        documents: documents,
        schedules: clean.schedules
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
      if (text_(masterValue_(row, 'Status Permohonan')) !== 'READY') {
        throw new Error('Hanya permohonan berstatus Siap Diproses yang dapat ditandai selesai.');
      }
      const progress = summarizeDocumentWorkflow_(getDocumentsByRequest_(id), 'READY');
      if (!progress.total || progress.status !== 'DRAFTED') {
        throw new Error('Seluruh dokumen harus berhasil dibuat dan seluruh draft email harus tersedia sebelum permohonan ditandai selesai.');
      }

      setMaster_(row, 'Status Permohonan', 'ARCHIVED');
      setMaster_(row, 'Diubah Oleh', user.email);
      setMaster_(row, 'Diubah Pada', new Date());
      setMaster_(row, 'Revision', currentRevision + 1);
      sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
      syncTravelDataInternal_(id);
      clearAppCache_();
      logAudit_('ARCHIVE_REQUEST', id, true, {});
      return { ok: true };
    });
  }

  function activateRequestInternal_(requestId, revision, user) {
    const id = text_(requestId);
    const sheet = getSheet_('MASTER');
    const rowNumber = findRowById_(sheet, id, 1);
    if (!rowNumber) throw new Error('Permohonan tidak ditemukan: ' + id);

    const row = sheet.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).getValues()[0];
    const request = masterRowToDto_(row);
    if (request.status === 'ARCHIVED') {
      throw new Error('Permohonan selesai bersifat hanya-baca.');
    }
    if (request.status === 'READY') {
      return { ok: true, reused: true, id: id, revision: request.revision };
    }
    if (revision != null && Number(revision) !== Number(request.revision)) {
      throw new Error('Revision tidak cocok. Muat ulang data sebelum memproses permohonan.');
    }

    const clean = normalizeRequestPayload_(Object.assign({}, request, {
      status: 'READY',
      documents: getDocumentsByRequest_(id),
      schedules: getSchedulesByRequest_(id, row),
      employees: getEmployeesByRequest_(id)
    }));
    validateRequestPayload_(clean);

    const nextRevision = Number(request.revision || 0) + 1;
    setMaster_(row, 'Status Permohonan', 'READY');
    setMaster_(row, 'Diubah Oleh', user.email);
    setMaster_(row, 'Diubah Pada', new Date());
    setMaster_(row, 'Revision', nextRevision);
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
    syncTravelDataInternal_(id);
    clearAppCache_();
    logAudit_('ACTIVATE_REQUEST', id, true, { revision: nextRevision });
    return { ok: true, reused: false, id: id, revision: nextRevision };
  }

  function getReferenceData() {
    assertAuthorized_();
    return serializeValue_(getReferenceDataInternal_());
  }

  function addReferenceCC(payload) {
    const user = assertAuthorized_();
    assertCanWrite_(user);

    const unit = text_(payload.unit).trim();
    const role = text_(payload.role).trim();
    const email = text_(payload.email).trim().toLowerCase();

    if (!unit || !role || !email) {
      throw new Error('Semua field wajib diisi.');
    }
    const emails = emailList_(email);
    if (!emails.length) {
      throw new Error('Format email tidak valid.');
    }

    return withScriptLock_(function() {
      const sheet = getSheet_('CC');
      appendDataRow_(sheet, [unit, role, emails[0]]);
      clearAppCache_();
      logAudit_('ADD_REFERENCE_CC', '', true, { unit: unit, role: role, email: emails[0] });
      return serializeValue_(getReferenceDataInternal_());
    });
  }


  function getReferenceDataInternal_() {
    const cached = getJsonCache_('references');
    if (cached) return cached;

    const ccSheet = getSheet_('CC');
    const lastRow = lastNonEmptyRowInColumn_(ccSheet, 1);
    const cc = lastRow < 2 ? [] :
      ccSheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
        .filter(function(row) { return row[1] && row[2]; })
        .map(function(row) {
          return { unit: row[0], role: row[1], email: row[2] };
        });
    return putJsonCache_('references', {
      cc: cc,
      signers: getSignatureConfigsInternal_()
    });
  }

  function getSignatureConfigsInternal_() {
    const cached = getJsonCache_('signers');
    if (cached) return cached;

    const sheet = getSheet_('SIGNATURE', false);
    if (!sheet) return [];

    const lastRow = lastNonEmptyRowInColumn_(sheet, 1);
    if (lastRow < 2) return [];

    const signers = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
      .map(function(row) {
        return {
          name: text_(row[0], 250),
          role: text_(row[1], 250),
          nik: text_(row[2], 100)
        };
      })
      .filter(function(item) {
        return item.name || item.role || item.nik;
      });
    return putJsonCache_('signers', signers);
  }

  function getEmployeeCatalogInternal_() {
    const cached = getJsonCache_('employeeCatalog');
    if (cached) return cached;

    const rows = readDataRows_(getSheet_('EMPLOYEES'), EMPLOYEE_HEADERS.length);
    const seen = {};
    const units = [];
    const roles = [];
    const ranks = [];
    const categories = [];

    function getIndex(arr, val) {
      const clean = text_(val);
      const idx = arr.indexOf(clean);
      if (idx !== -1) return idx;
      arr.push(clean);
      return arr.length - 1;
    }

    const items = [];
    rows.forEach(function(row) {
      const name = text_(row[1], 250);
      const identifier = text_(row[2], 100);
      const email = text_(row[5], 250);
      if (!name && !identifier && !email) return;

      const unit = canonicalOrganizationUnit_(row[4]);
      const role = text_(row[3], 200);
      const rank = text_(row[6], 200);
      const category = text_(row[7], 200);

      const key = [identifier, email.toLowerCase(), name.toLowerCase(), unit.toLowerCase()]
        .filter(Boolean)
        .join('|');
      if (seen[key]) return;
      seen[key] = true;

      items.push([
        name,
        identifier,
        getIndex(roles, role),
        getIndex(units, unit),
        email,
        getIndex(ranks, rank),
        getIndex(categories, category)
      ]);
    });

    items.sort(function(a, b) {
      return a[0].localeCompare(b[0]) || a[1].localeCompare(b[1]);
    });

    const payload = {
      units: units,
      roles: roles,
      ranks: ranks,
      categories: categories,
      items: items
    };

    return putJsonCache_('employeeCatalog', payload);
  }

  function buildDashboardSummary_() {
    const documentsByRequest = buildDocumentsByRequestMap_();
    const requests = readDataRows_(getSheet_('MASTER'), MASTER_HEADERS.length)
      .map(function(row) {
        return enrichRequestWithDocuments_(masterRowToDto_(row), documentsByRequest);
      })
      .filter(function(item) {
        return item.id && item.status !== 'ARCHIVED';
      });
    const processableRequestIds = requests.reduce(function(result, item) {
      if (item.status === 'READY') result[item.id] = true;
      return result;
    }, {});
    const documents = readDataRows_(getSheet_('DOCUMENTS'), DOCUMENT_HEADERS.length)
      .map(documentRowToDto_)
      .filter(function(item) { return processableRequestIds[item.requestId]; });
    const now = new Date();
    const next30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);

    const pendingFollowUp = requests.filter(function(req) {
      if (req.status !== 'READY') return false;
      const progress = req.documentProgress;
      return !progress || (progress.status !== 'DRAFTED' && progress.status !== 'COMPLETE');
    }).length;

    return {
      total: requests.length,
      drafts: requests.filter(function(item) { return item.status === 'DRAFT'; }).length,
      ready: requests.filter(function(item) { return item.status === 'READY'; }).length,
      pendingFollowUp: pendingFollowUp,
      documentsPending: documents.filter(function(item) {
        return item.status === 'PENDING';
      }).length,
      documentsFailed: documents.filter(function(item) {
        return item.status === 'ERROR';
      }).length,
      emailsPending: documents.filter(function(item) {
        return item.status === 'GENERATED' && item.emailStatus !== 'DRAFTED';
      }).length,
      emailsDrafted: documents.filter(function(item) {
        return item.emailStatus === 'DRAFTED';
      }).length,
      upcoming: requests.filter(function(item) {
        const startIso = normalizeIsoDateString_(item.startDate);
        if (!startIso) return false;
        const date = parseIsoDate_(startIso);
        return date >= now && date <= next30;
      }).length
    };
  }

  function normalizeIsoDateString_(value) {
    const raw = text_(value);
    if (!raw) return '';
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;

    const date = value instanceof Date ? value : new Date(raw);
    if (isNaN(date.getTime())) return '';
    return formatDate_(date, 'yyyy-MM-dd');
  }

  function masterRowToDto_(row) {
    // Parse manual To/CC selections stored as MANUAL_JSON prefix in Keterangan Email
    var rawNote = String(masterValue_(row, 'Keterangan Email') || '');
    var manualTo = [];
    var manualCc = [];
    var emailNote = rawNote;
    if (rawNote.indexOf('MANUAL_JSON:') === 0) {
      try {
        var newlinePos = rawNote.indexOf('\n');
        var jsonStr = newlinePos === -1 ? rawNote.slice(12) : rawNote.slice(12, newlinePos);
        var parsed = JSON.parse(jsonStr);
        manualTo = Array.isArray(parsed.to) ? parsed.to : [];
        manualCc = Array.isArray(parsed.cc) ? parsed.cc : [];
        emailNote = newlinePos === -1 ? '' : rawNote.slice(newlinePos + 1);
      } catch (e) {
        // Parsing failed, treat raw as note
      }
    }
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
      startDate: normalizeIsoDateString_(masterValue_(row, 'Tanggal Mulai ISO')),
      endDate: normalizeIsoDateString_(masterValue_(row, 'Tanggal Selesai ISO')),
      activityTime: masterValue_(row, 'Waktu Kegiatan'),
      activityPlace: masterValue_(row, 'Tempat Kegiatan'),
      signerName: masterValue_(row, 'Nama Penandatangan'),
      signerId: masterValue_(row, 'NIK Penandatangan'),
      signerRole: masterValue_(row, 'Jabatan Penandatangan'),
      honor: masterValue_(row, 'Honor') || 'Tidak',
      travel: masterValue_(row, 'Perjalanan Dinas') || 'Tidak',
      emailTo: masterValue_(row, 'Email To'),
      emailCc: masterValue_(row, 'Email CC'),
      emailNote: emailNote,
      emailStatus: masterValue_(row, 'Email Status'),
      status: masterValue_(row, 'Status Permohonan') || 'DRAFT',
      createdBy: masterValue_(row, 'Dibuat Oleh'),
      createdAt: serializeValue_(masterValue_(row, 'Dibuat Pada')),
      updatedBy: masterValue_(row, 'Diubah Oleh'),
      updatedAt: serializeValue_(masterValue_(row, 'Diubah Pada')),
      clientToken: masterValue_(row, 'Client Token'),
      revision: Number(masterValue_(row, 'Revision') || 0),
      manualTo: manualTo,
      manualCc: manualCc
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
      updatedAt: serializeValue_(row[16]),
      emailTo: row[17] || '',
      emailCc: row[18] || '',
      emailBcc: row[19] || ''
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
          unit: canonicalOrganizationUnit_(row[4]),
          email: row[5],
          rank: row[6],
          category: row[7],
          participantKey: row[8]
        };
      });
  }

  function getSchedulesByRequest_(requestId, masterRow) {
    const sheet = getSheet_('SCHEDULES', false);
    const rows = sheet
      ? readDataRows_(sheet, SCHEDULE_HEADERS.length)
        .filter(function(row) { return text_(row[1]) === requestId; })
      : [];
    const schedules = rows.map(scheduleRowToDto_).sort(function(a, b) {
      return Number(a.sequence || 0) - Number(b.sequence || 0);
    });
    if (schedules.length) return schedules;

    const row = masterRow || [];
    const startDate = normalizeIsoDateString_(masterValue_(row, 'Tanggal Mulai ISO'));
    if (!startDate) return [];
    const legacyTime = parseTimeRange_(masterValue_(row, 'Waktu Kegiatan'));
    const legacySchedule = {
      id: '',
      requestId: requestId,
      startDate: startDate,
      endDate: normalizeIsoDateString_(masterValue_(row, 'Tanggal Selesai ISO')) || startDate,
      startTime: legacyTime.startTime,
      endTime: legacyTime.endTime,
      place: text_(masterValue_(row, 'Tempat Kegiatan')),
      sequence: 1
    };
    legacySchedule.label = formatScheduleItem_(legacySchedule);
    return [legacySchedule];
  }

  function scheduleRowToDto_(row) {
    const schedule = {
      id: row[0],
      requestId: row[1],
      startDate: normalizeIsoDateString_(row[2]),
      endDate: normalizeIsoDateString_(row[3]) || normalizeIsoDateString_(row[2]),
      startTime: normalizeTimeValue_(row[4]),
      endTime: normalizeTimeValue_(row[5]),
      place: text_(row[6]),
      sequence: Number(row[7] || 0)
    };
    schedule.label = formatScheduleItem_(schedule);
    return schedule;
  }

  function replaceDocumentsForRequest_(requestId, documents, revision, preserveGenerated) {
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
      const reusable = old &&
        preserveGenerated &&
        text_(old[2]) === document.type &&
        text_(old[5]) === document.number &&
        text_(old[6]) === templateKey;
      row[0] = old ? old[0] : 'DOC-' + Utilities.getUuid().slice(0, 12).toUpperCase();
      row[1] = requestId;
      row[2] = document.type;
      row[3] = document.speakerSubtype;
      row[4] = document.speakerStatus;
      row[5] = document.number;
      row[6] = templateKey;
      row[7] = reusable
        ? (old[7] || 'PENDING')
        : 'PENDING';
      if (!reusable) {
        for (let i = 8; i <= 13; i++) row[i] = '';
      }
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

  function syncDocumentRecipientsForRequest_(requestId, request, employees) {
    const sheet = getSheet_('DOCUMENTS');
    const rows = readDataRows_(sheet, DOCUMENT_HEADERS.length);
    let changed = false;

    rows.forEach(function(row) {
      if (text_(row[1]) !== requestId) return;
      const stored = documentRowToDto_(row);
      const routing = computeEmailRouting_({
        activityType: request.activityType,
        documents: [stored],
        partnerEmail: request.partnerEmail,
        partnerName: request.partnerName,
        faculties: request.faculties,
        speakerSubtype: stored.speakerSubtype,
        speakerStatus: stored.speakerStatus,
        manualTo: request.manualTo || [],
        manualCc: request.manualCc || []
      }, employees, stored.type);

      const nextTo = routing.to.join(', ');
      const nextCc = routing.cc.join(', ');
      const nextBcc = routing.bcc.join(', ');
      if (text_(row[17]) !== nextTo || text_(row[18]) !== nextCc || text_(row[19]) !== nextBcc) {
        row[12] = '';
        row[13] = 'PENDING';
      }
      row[17] = nextTo;
      row[18] = nextCc;
      row[19] = nextBcc;
      changed = true;
    });

    if (changed) rewriteDataRows_(sheet, rows, DOCUMENT_HEADERS.length);
  }

  function replaceSchedulesForRequest_(requestId, schedules) {
    const sheet = getSheet_('SCHEDULES');
    const untouched = readDataRows_(sheet, SCHEDULE_HEADERS.length)
      .filter(function(row) { return text_(row[1]) !== requestId; });
    const additions = schedules.map(function(schedule, index) {
      return [
        schedule.id || 'SCH-' + Utilities.getUuid().slice(0, 12).toUpperCase(),
        requestId,
        schedule.startDate ? parseIsoDate_(schedule.startDate) : '',
        schedule.startDate ? parseIsoDate_(schedule.endDate || schedule.startDate) : '',
        schedule.startTime,
        schedule.endTime,
        schedule.place,
        index + 1
      ];
    });
    rewriteDataRows_(sheet, untouched.concat(additions), SCHEDULE_HEADERS.length);
  }

  function clearGeneratedStateColumns_(row) {
    setMaster_(row, 'Edit Surat', '');
    setMaster_(row, 'Download PDF Surat', '');
    setMaster_(row, 'Email Status', '');
  }

  function requestGenerationFingerprint_(request) {
    const clean = request || {};
    return JSON.stringify({
      activityType: clean.activityType,
      speakerSubtype: clean.speakerSubtype,
      speakerStatus: clean.speakerStatus,
      faculties: clean.faculties || [],
      incomingNumber: clean.incomingNumber,
      incomingDate: clean.incomingDate,
      incomingSender: clean.incomingSender,
      incomingSubject: clean.incomingSubject,
      activityName: clean.activityName,
      partnerName: clean.partnerName,
      partnerAddress: clean.partnerAddress,
      partnerEmail: clean.partnerEmail,
      letterDate: clean.letterDate,
      signerName: clean.signerName,
      signerId: clean.signerId,
      signerRole: clean.signerRole,
      honor: clean.honor,
      travel: clean.travel,
      manualTo: clean.manualTo || [],
      manualCc: clean.manualCc || [],
      documents: (clean.documents || []).map(function(item) {
        return [item.type, item.number, item.speakerSubtype, item.speakerStatus];
      }),
      schedules: (clean.schedules || []).map(function(item) {
        return [item.startDate, item.endDate, item.startTime, item.endTime, item.place];
      }),
      employees: (clean.employees || []).map(function(item) {
        return [
          item.name, item.identifier, item.role, item.unit, item.email,
          item.rank, item.category
        ];
      })
    });
  }

  function markGeneratedFilesSuperseded_(requestId) {
    const sheet = getSheet_('FILES', false);
    if (!sheet) return;
    const rows = readDataRows_(sheet, GENERATED_FILE_HEADERS.length);
    let changed = false;
    rows.forEach(function(row) {
      if (text_(row[0]) === requestId && text_(row[8]) === 'ACTIVE') {
        row[8] = 'SUPERSEDED';
        changed = true;
      }
    });
    if (changed) rewriteDataRows_(sheet, rows, GENERATED_FILE_HEADERS.length);
  }

  function normalizeRequestPayload_(payload) {
    const activityType = text_(payload.activityType);
    let speakerSubtype = text_(payload.speakerSubtype);
    let speakerStatus = text_(payload.speakerStatus);
    if (activityType !== 'Penugasan Narasumber') {
      speakerSubtype = '';
      speakerStatus = '';
    }
    const documents = (payload.documents || []).map(function(document) {
      return {
        id: text_(document.id),
        activityType: activityType,
        type: text_(document.type),
        speakerSubtype: activityType === 'Penugasan Narasumber'
          ? text_(document.speakerSubtype || speakerSubtype)
          : '',
        speakerStatus: activityType === 'Penugasan Narasumber'
          ? text_(document.speakerStatus || speakerStatus)
          : '',
        number: text_(document.number, 150),
        emailTo: emailList_(document.emailTo).join(', '),
        emailCc: emailList_(document.emailCc).join(', '),
        emailBcc: emailList_(document.emailBcc).join(', '),
        recipientsProvided: Object.prototype.hasOwnProperty.call(document, 'emailTo') ||
          Object.prototype.hasOwnProperty.call(document, 'emailCc') ||
          Object.prototype.hasOwnProperty.call(document, 'emailBcc')
      };
    });
    const employees = (payload.employees || []).map(function(employee) {
      return {
        name: text_(employee.name, 200),
        identifier: text_(employee.identifier, 100),
        role: text_(employee.role, 200),
        unit: canonicalOrganizationUnit_(employee.unit),
        email: emailList_(employee.email).join(', '),
        rank: text_(employee.rank, 200),
        category: text_(employee.category, 200),
        participantKey: text_(employee.participantKey) || buildParticipantKey_(payload.id, employee)
      };
    }).filter(function(employee) {
      return employee.name || employee.identifier || employee.email || employee.role ||
        employee.unit || employee.rank || employee.category;
    });
    let schedules = (payload.schedules || []).map(function(schedule, index) {
      return {
        id: text_(schedule.id),
        startDate: text_(schedule.startDate),
        endDate: text_(schedule.endDate) || text_(schedule.startDate),
        startTime: text_(schedule.startTime, 20),
        endTime: text_(schedule.endTime, 20),
        place: text_(schedule.place, 500),
        sequence: index + 1
      };
    }).filter(function(schedule) {
      return schedule.startDate || schedule.startTime || schedule.endTime || schedule.place;
    });
    if (!schedules.length && text_(payload.startDate)) {
      schedules = [{
        id: '',
        startDate: text_(payload.startDate),
        endDate: text_(payload.endDate) || text_(payload.startDate),
        startTime: '',
        endTime: '',
        place: text_(payload.activityPlace, 500),
        sequence: 1
      }];
    }
    const scheduleSummary = schedules.length ? buildScheduleSummary_(schedules) : {
      startDate: '',
      endDate: '',
      timeDisplay: text_(payload.activityTime, 100),
      placeDisplay: text_(payload.activityPlace, 500)
    };
    const manualTo = emailList_(Array.isArray(payload.manualTo) ? payload.manualTo : []);
    const manualCc = emailList_(Array.isArray(payload.manualCc) ? payload.manualCc : [])
      .filter(function(email) {
        return manualTo.indexOf(email) === -1;
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
      startDate: scheduleSummary.startDate,
      endDate: scheduleSummary.endDate,
      activityTime: scheduleSummary.timeDisplay,
      activityPlace: scheduleSummary.placeDisplay,
      signerName: text_(payload.signerName, 250),
      signerId: text_(payload.signerId, 100),
      signerRole: text_(payload.signerRole, 250),
      honor: text_(payload.honor) === 'Ya' ? 'Ya' : 'Tidak',
      travel: text_(payload.travel) === 'Ya' ? 'Ya' : 'Tidak',
      documents: documents,
      employees: employees,
      schedules: schedules,
      manualTo: manualTo,
      manualCc: manualCc
    };
  }

  function validateRequestPayload_(payload) {
    const errors = [];
    const isReady = payload.status === 'READY';
    const schedules = payload.schedules || (payload.startDate ? [{
      startDate: payload.startDate,
      endDate: payload.endDate || payload.startDate,
      startTime: '',
      endTime: '',
      place: payload.activityPlace || ''
    }] : []);
    const documents = payload.documents || [];
    const employees = payload.employees || [];
    if (APP_CONFIG.ACTIVITY_TYPES.indexOf(payload.activityType) === -1) {
      errors.push('Tipe kegiatan tidak valid.');
    }
    if (isReady && !payload.activityName) errors.push('Nama kegiatan wajib diisi.');
    if (isReady && !payload.partnerName) errors.push('Nama mitra wajib diisi.');
    if (isReady && !schedules.length) errors.push('Minimal satu sesi jadwal wajib diisi.');

    schedules.forEach(function(schedule, index) {
      if (!schedule.startDate) {
        if (isReady) errors.push('Tanggal sesi ke-' + (index + 1) + ' wajib diisi.');
        return;
      }
      const start = parseIsoDate_(schedule.startDate);
      const end = parseIsoDate_(schedule.endDate || schedule.startDate);
      if (end.getTime() < start.getTime()) {
        errors.push('Tanggal selesai sesi ke-' + (index + 1) + ' tidak boleh sebelum tanggal mulai.');
      }
      if (schedule.startTime && schedule.endTime && schedule.endTime < schedule.startTime) {
        errors.push('Waktu selesai sesi ke-' + (index + 1) + ' tidak boleh sebelum waktu mulai.');
      }
      if (isReady && !schedule.place) {
        errors.push('Tempat sesi ke-' + (index + 1) + ' wajib diisi.');
      }
    });
    if (payload.incomingDate) parseIsoDate_(payload.incomingDate);
    if (payload.letterDate) parseIsoDate_(payload.letterDate);

    const seenDocumentTypes = {};
    documents.forEach(function(document) {
      if (seenDocumentTypes[document.type]) {
        errors.push('Jenis dokumen tidak boleh dipilih lebih dari sekali: ' + document.type);
      }
      seenDocumentTypes[document.type] = true;

      const condition = (function() {
        if (payload.activityType === 'Penugasan Narasumber') {
          if (document.type === 'Surat Tugas') return document.speakerSubtype || '-';
          if (document.type === 'Surat Permohonan Narasumber kepada Dekan') return document.speakerStatus || '-';
        }
        return '-';
      })();

      const mapping = findOfficialMapping_(payload.activityType, document.type, condition);
      if (!mapping) {
        errors.push(
          'Kombinasi Tipe Kegiatan (' + payload.activityType + '), Sub-Tipe (' + document.type + 
          '), dan Kondisi Tambahan (' + condition + ') tidak ada pada mapping resmi.'
        );
      }

      if (isReady && !document.number) {
        errors.push('Nomor surat wajib untuk status siap: ' + document.type);
      }
    });

    const hasRequestLetter = documents.find(function(document) {
      return document.type === 'Surat Permohonan Narasumber kepada Dekan';
    });
    const needsPeople = documents.some(function(document) {
      return document.type === 'Surat Tugas';
    }) || Boolean(hasRequestLetter);
    if (isReady && needsPeople && !employees.length) {
      errors.push('Minimal satu pegawai/narasumber wajib diisi.');
    }
    if (isReady) {
      employees.forEach(function(employee, index) {
        const missing = [
          ['Nama', employee.name],
          ['NIP/NPM', employee.identifier],
          ['Email', employee.email],
          ['Jabatan', employee.role],
          ['Fakultas', employee.unit]
        ].filter(function(item) {
          return !item[1];
        }).map(function(item) {
          return item[0];
        });
        if (missing.length) {
          errors.push(
            'Data orang ke-' + (index + 1) + ' belum lengkap: ' + missing.join(', ') + '.'
          );
        }
        if (employee.unit && APP_CONFIG.FACULTIES.indexOf(canonicalOrganizationUnit_(employee.unit)) === -1) {
          errors.push(
            'Unit/Fakultas orang ke-' + (index + 1) + ' tidak dikenali: ' + employee.unit + '.'
          );
        }
      });
    }

    if (errors.length) throw new Error(errors.join('\n'));
  }

  function allowedDocumentsFor_(activityType) {
    const list = [];
    OFFICIAL_LETTER_MAPPING.forEach(function(item) {
      if (item.activityType === activityType) {
        if (list.indexOf(item.subType) === -1) {
          list.push(item.subType);
        }
      }
    });
    return list;
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

    // Store manual To/CC selections as structured JSON prefix in Keterangan Email
    // so they can be restored when the form is re-opened for editing.
    const manualMeta = 'MANUAL_JSON:' + JSON.stringify({
      to: request.manualTo || [],
      cc: request.manualCc || []
    });
    const noteText = routing.notes.length ? routing.notes.join(' | ') : '';
    setMaster_(row, 'Keterangan Email', manualMeta + (noteText ? '\n' + noteText : ''));
  }

  function computeEmailRouting_(request, employees, documentType) {
    const references = getReferenceDataInternal_().cc;
    const to = [];
    const toRoles = [];
    const cc = [];
    const ccRoles = [];
    const bcc = [];
    const notes = [];
    const types = documentType
      ? [documentType]
      : request.documents.map(function(document) { return document.type; });

    const configs = getTemplateConfigsInternal_();

    function addByRole(targetEmails, targetRoles, role, unit) {
      function lookupKey(value) {
        return String(value || '').toLowerCase()
          .replace(/&/g, ' dan ')
          .replace(/[^a-z0-9]+/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
      }
      const canonicalUnit = canonicalOrganizationUnit_(unit);
      const match = references.find(function(item) {
        return lookupKey(item.role) === lookupKey(role) &&
          (!canonicalUnit ||
            lookupKey(canonicalOrganizationUnit_(item.unit)) === lookupKey(canonicalUnit));
      });
      if (match) {
        targetEmails.push(match.email);
        targetRoles.push(match.role);
      } else {
        notes.push('Master CC tidak menemukan: ' + role);
      }
    }

    function addEmployees(targetEmails, targetRoles) {
      targetEmails = targetEmails || to;
      targetRoles = targetRoles || toRoles;
      employees.forEach(function(employee) {
        emailList_(employee.email).forEach(function(email) { targetEmails.push(email); });
        if (employee.name) targetRoles.push(employee.name);
      });
    }

    function addDeansForFaculties(faculties) {
      uniqueTextList_(faculties || []).forEach(function(faculty) {
        addByRole(to, toRoles, 'Dekan ' + faculty, faculty);
      });
    }

    function employeeFaculties() {
      return uniqueTextList_((employees || []).map(function(employee) {
        return canonicalOrganizationUnit_(employee.unit);
      }).filter(function(unit) {
        return APP_CONFIG.FACULTIES.indexOf(unit) !== -1 && unit.indexOf('Fakultas ') === 0;
      }));
    }

    function addManualEmail(targetEmails, targetRoles, email) {
      if (!email || targetEmails.indexOf(email) !== -1) return;
      const normalizedEmail = emailList_(email)[0];
      const ref = references.find(function(item) {
        return String(item.email || '').trim().toLowerCase() === normalizedEmail;
      });
      if (!ref) {
        notes.push('Penerima manual tidak ditemukan di Master CC: ' + email);
        return;
      }
      targetEmails.push(normalizedEmail);
      targetRoles.push(ref.role || ref.unit || normalizedEmail);
    }

    function resolveConfiguredRecipients(configuredText, targetEmails, targetRoles) {
      if (!configuredText) return;
      let remaining = String(configuredText);
      const normalized = remaining.trim().toLowerCase();
      if (normalized.indexOf('email petugas/narasumber') !== -1) {
        addEmployees(targetEmails, targetRoles);
        remaining = remaining.replace(/email petugas\/narasumber/ig, '');
      }
      if (normalized.indexOf('email mitra') !== -1) {
        emailList_(request.partnerEmail).forEach(function(email) { targetEmails.push(email); });
        if (request.partnerName) targetRoles.push(request.partnerName);
        remaining = remaining.replace(/email mitra/ig, '');
      }

      references.slice().sort(function(a, b) {
        return String(b.role || '').length - String(a.role || '').length;
      }).forEach(function(reference) {
        const role = String(reference.role || '').trim();
        if (!role || remaining.toLowerCase().indexOf(role.toLowerCase()) === -1) return;
        addByRole(targetEmails, targetRoles, role, reference.unit);
        remaining = remaining.replace(new RegExp(escapeRegex_(role), 'ig'), '');
      });

      const parts = remaining.split(/[;\n]+/).map(function(p) {
        return p.replace(/^[,\s]+|[,\s]+$/g, '').trim();
      }).filter(Boolean);
      parts.forEach(function(part) {
        if (part.indexOf('@') !== -1) {
          targetEmails.push(part.toLowerCase());
          targetRoles.push(part);
        } else {
          let role = part;
          let unit = '';
          if (part.indexOf('|') !== -1) {
            const split = part.split('|');
            role = split[0].trim();
            unit = split[1].trim();
          }
          addByRole(targetEmails, targetRoles, role, unit);
        }
      });
    }

    types.forEach(function(type) {
      let templateKey = null;
      try {
        const docObj = (request.documents || []).find(function(d) { return d.type === type; }) || {};
        templateKey = resolveTemplateKey_({
          type: type,
          activityType: request.activityType,
          speakerSubtype: docObj.speakerSubtype || request.speakerSubtype || '',
          speakerStatus: docObj.speakerStatus || request.speakerStatus || ''
        });
      } catch (e) {}

      const cfg = (templateKey
        ? configs.find(function(c) { return c.key === templateKey && c.active; })
        : null) || configs.find(function(c) {
        return c.type === type && c.active;
      });

      if (cfg) {
        if (type !== 'Surat Permohonan Narasumber kepada Dekan') {
          resolveConfiguredRecipients(cfg.defaultTo, to, toRoles);
        }
        resolveConfiguredRecipients(cfg.defaultCc, cc, ccRoles);
        resolveConfiguredRecipients(cfg.defaultBcc, bcc, []);
      } else {
        // Fallback rules
        if (type === 'Surat Balasan Campus Visit') {
          emailList_(request.partnerEmail).forEach(function(email) { to.push(email); });
          toRoles.push(request.partnerName);
          // addByRole(cc, ccRoles, 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', 'Rektorat');
        } else if (type === 'Surat Rekomendasi Campus Visit - SU') {
          addByRole(to, toRoles, 'Sekretaris Universitas', 'Rektorat');
          // addByRole(cc, ccRoles, 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', 'Rektorat');
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

          /*
          [
            ['Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', 'Rektorat'],
            ['Manajer Aset dan Sarana Prasarana', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
            ['Manajer Kemahasiswaan', 'Direktorat Kemahasiswaan'],
            ['Koordinator Kebersihan, Keamanan dan Ketertiban', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
            ['Koordinator Kelas dan Fasilitas Pendukung', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
            ['Koordinator Pemeliharaan Kelas', 'Direktorat Manajemen Aset, Keuangan, dan Sarana Prasarana'],
            ['Koordinator Administrasi Fakultas Ekonomi', 'Fakultas Ekonomi'],
            ['Koordinator Administrasi Fakultas Hukum', 'Fakultas Hukum'],
            ['Koordinator Administrasi Fakultas Ilmu Sosial dan Ilmu Politik', 'Fakultas Ilmu Sosial dan Ilmu Politik'],
            ['Koordinator Administrasi Fakultas Sains', 'Fakultas Sains'],
            ['Koordinator Administrasi Fakultas Teknik', 'Fakultas Teknik'],
            ['Koordinator Administrasi Fakultas Teknologi Rekayasa', 'Fakultas Teknologi Rekayasa'],
            ['Koordinator Administrasi Fakultas Vokasi', 'Fakultas Vokasi']
          ].forEach(function(pair) { addByRole(cc, ccRoles, pair[0], pair[1]); });
          */
        }
      }

      if (type === 'Surat Balasan Campus Visit' && to.length === 0) {
        emailList_(request.partnerEmail).forEach(function(email) { to.push(email); });
        toRoles.push(request.partnerName);
      }

      if (type === 'Surat Permohonan Narasumber kepada Dekan') {
        const sourceFaculties = employeeFaculties();
        addDeansForFaculties(sourceFaculties.length ? sourceFaculties : request.faculties);
      } else if (type === 'Surat Tugas') {
        addEmployees();
      }
    });

    (request.manualCc || []).forEach(function(email) {
      addManualEmail(cc, ccRoles, email);
    });

    const normalizedTo = emailList_(to);
    const normalizedCc = emailList_(cc).filter(function(email) {
      return normalizedTo.indexOf(email) === -1;
    });
    const overlappingCcRoles = references.filter(function(item) {
      return normalizedTo.indexOf(String(item.email || '').trim().toLowerCase()) !== -1;
    }).map(function(item) {
      return item.role || item.unit || item.email;
    });

    return {
      to: normalizedTo,
      toRoles: uniqueTextList_(toRoles),
      cc: normalizedCc,
      ccRoles: uniqueTextList_(ccRoles).filter(function(role) {
        return overlappingCcRoles.indexOf(role) === -1;
      }),
      bcc: uniqueTextList_(bcc),
      notes: uniqueTextList_(notes)
    };
  }

  function findOfficialMapping_(activityType, subType, condition) {
    const normActivity = String(activityType || '').trim();
    const normSubType = String(subType || '').trim();
    const normCondition = String(condition || '').trim() || '-';
    
    return OFFICIAL_LETTER_MAPPING.find(function(item) {
      return item.activityType === normActivity &&
             item.subType === normSubType &&
             item.condition === normCondition;
    }) || null;
  }

  function resolveTemplateKey_(document) {
    const descriptor = normalizeDocumentDescriptor_(document);
    
    const condition = (function() {
      if (descriptor.activityType === 'Penugasan Narasumber') {
        if (descriptor.type === 'Surat Tugas') return descriptor.speakerSubtype || '-';
        if (descriptor.type === 'Surat Permohonan Narasumber kepada Dekan') return descriptor.speakerStatus || '-';
      }
      return '-';
    })();

    const mapping = findOfficialMapping_(descriptor.activityType, descriptor.type, condition);
    if (mapping) return mapping.templateKey;

    throw new Error(
      'Template tidak tersedia untuk kombinasi dokumen: ' +
      [descriptor.activityType, descriptor.type, condition]
        .map(function(value) { return text_(value) || '-'; })
        .join(' | ')
    );
  }

  function normalizeDocumentDescriptor_(document) {
    const descriptor = {
      activityType: text_((document || {}).activityType),
      type: text_((document || {}).type),
      speakerSubtype: text_((document || {}).speakerSubtype),
      speakerStatus: text_((document || {}).speakerStatus)
    };

    const activityMap = {
      'edu fair': 'Edu Fair',
      'campus visit': 'Campus Visit',
      'penugasan narasumber': 'Penugasan Narasumber',
      'narasumber': 'Penugasan Narasumber'
    };
    const typeMap = {
      'surat tugas': 'Surat Tugas',
      'surat balasan campus visit': 'Surat Balasan Campus Visit',
      'surat rekomendasi campus visit su': 'Surat Rekomendasi Campus Visit - SU',
      'surat rekomendasi campus visit - su': 'Surat Rekomendasi Campus Visit - SU',
      'surat izin pimpinan campus visit': 'Surat izin pimpinan - Campus Visit',
      'surat izin pimpinan - campus visit': 'Surat izin pimpinan - Campus Visit',
      'surat ijin pimpinan campus visit': 'Surat izin pimpinan - Campus Visit',
      'surat ijin pimpinan - campus visit': 'Surat izin pimpinan - Campus Visit',
      'surat permohonan narasumber kepada dekan': 'Surat Permohonan Narasumber kepada Dekan'
    };
    const subtypeMap = {
      'workshop': 'Workshop',
      'promosi': 'Promosi'
    };
    const statusMap = {
      'dicarikan': 'Dicarikan',
      'tidak dicarikan': 'Tidak Dicarikan'
    };

    const canonical = function(value) {
      return text_(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
    };

    const normalized = {
      activityType: activityMap[canonical(descriptor.activityType)] || descriptor.activityType,
      type: typeMap[canonical(descriptor.type)] || descriptor.type,
      speakerSubtype: subtypeMap[canonical(descriptor.speakerSubtype)] || descriptor.speakerSubtype,
      speakerStatus: statusMap[canonical(descriptor.speakerStatus)] || descriptor.speakerStatus
    };

    if (normalized.activityType === 'Penugasan Narasumber' &&
        normalized.type === 'Surat Permohonan Narasumber kepada Dekan' &&
        !normalized.speakerStatus) {
      normalized.speakerStatus = 'Tidak Dicarikan';
    }

    return normalized;
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
    const lastRow = Math.max(Number(sheet.getLastRow()) || 0, 1);
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, width).getValues()
      .filter(function(row) {
        return row.some(function(value) { return value !== ''; });
      });
  }

  function rewriteDataRows_(sheet, rows, width) {
    const oldRows = Math.max((Number(sheet.getLastRow()) || 1) - 1, 0);
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
    if (['ADMIN', 'OPERATOR'].indexOf((user || {}).role) === -1) {
      throw new Error('Hanya Admin atau Operator dapat mengubah data.');
    }
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

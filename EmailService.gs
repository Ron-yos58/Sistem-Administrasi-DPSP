function previewEmail(documentId, customRecipientsJson) {
  assertAuthorized_();
  const document = getDocumentById_(text_(documentId));
  let customRecipients = null;
  if (customRecipientsJson) {
    try {
      customRecipients = JSON.parse(customRecipientsJson);
    } catch (e) {
      console.error('Failed to parse customRecipientsJson: ' + e.message);
    }
  }
  return serializeValue_(buildEmailPreviewForDocument_(document, customRecipients));
}

function createEmailDraft(documentId, force, customRecipientsJson) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(documentId);

  return withScriptLock_(function() {
    try {
      const result = createEmailDraftInternal_(id, Boolean(force), user, customRecipientsJson);
      logAudit_('CREATE_EMAIL_DRAFT', result.requestId, true, {
        documentId: id,
        reused: result.reused
      });
      return serializeValue_(result);
    } catch (error) {
      logAudit_('CREATE_EMAIL_DRAFT', id, false, { error: error.message });
      throw error;
    }
  });
}


function processRequest(requestId, revision) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(requestId);

  return withScriptLock_(function() {
    const result = activateRequestInternal_(id, revision, user);
    return serializeValue_(result);
  });
}

function buildEmailPreviewInternal_(requestId, suppressErrors) {
  return getDocumentsByRequest_(requestId).map(function(document) {
    try {
      return buildEmailPreviewForDocument_(document);
    } catch (error) {
      if (!suppressErrors) throw error;
      return {
        documentId: document.id,
        type: document.type,
        error: error.message
      };
    }
  });
}

function buildEmailPreviewForDocument_(document, customRecipients) {
  const detail = getRequestDetailInternal_(document.requestId);
  const request = detail.request;
  let routing;

  if (customRecipients && (customRecipients.to || customRecipients.cc || customRecipients.bcc)) {
    const to = parseAndValidateEmailString_(customRecipients.to || '');
    const cc = parseAndValidateEmailString_(customRecipients.cc || '');
    const bcc = parseAndValidateEmailString_(customRecipients.bcc || '');

    saveFinalRecipientsToMaster_(document.requestId, to, cc, document.id);

    routing = {
      to: to,
      cc: cc,
      bcc: bcc,
      toRoles: customRecipients.toRoles ? customRecipients.toRoles.split(/[,;\n\s]+/).filter(Boolean) : to,
      ccRoles: customRecipients.ccRoles ? customRecipients.ccRoles.split(/[,;\n\s]+/).filter(Boolean) : cc,
      notes: []
    };
  } else {
    const to = parseAndValidateEmailString_(document.emailTo || '');
    const cc = parseAndValidateEmailString_(document.emailCc || '');
    const bcc = parseAndValidateEmailString_(document.emailBcc || '');
    routing = {
      to: to,
      cc: cc,
      bcc: bcc,
      toRoles: to,
      ccRoles: cc,
      notes: []
    };
  }

  if (!routing.to.length) {
    throw new Error('Penerima To belum tersedia. Periksa email pegawai, mitra, atau Master CC.');
  }
  if (routing.to.length + routing.cc.length > 50) {
    throw new Error('Jumlah penerima melebihi batas aman 50 alamat.');
  }

  const template = getEmailTemplate_(document.type, document.speakerStatus);
  const placeholders = buildEmailPlaceholders_(detail, document, routing);
  const subject = replaceTemplateTokens_(template.subject, placeholders, false);
  const isHtmlTemplate = isExplicitHtmlTemplate_(template.body);
  const renderedBody = replaceTemplateTokens_(template.body, placeholders, isHtmlTemplate);
  const htmlBody = isHtmlTemplate ? renderedBody : normalizeEmailHtml_(renderedBody);

  return {
    documentId: document.id,
    requestId: document.requestId,
    type: document.type,
    to: routing.to,
    cc: routing.cc,
    bcc: routing.bcc || [],
    toRoles: routing.toRoles,
    ccRoles: routing.ccRoles,
    notes: routing.notes,
    subject: subject,
    htmlBody: htmlBody,
    hasAttachment: Boolean(document.pdfId && driveFileExists_(document.pdfId)),
    attachmentUrl: document.pdfUrl || ''
  };
}

function normalizeEmailHtml_(html) {
  const source = String(html || '').trim();
  if (!source) return '';

  // Keep explicit HTML templates unchanged.
  if (isExplicitHtmlTemplate_(source)) return source;

  let text = source
    .replace(/\\n/g, '\n')
    .replace(/\r\n?/g, '\n')
    .trim();

  // Handle inline legacy templates where placeholders and salutations are stuck together.
  text = text
    .replace(/\s*Kepada\s+Yth\.?\s*:?[ \t]*/gi, 'Kepada Yth.:\n')
    .replace(/\s*Dengan hormat,?/gi, '\n\nDengan hormat,')
    .replace(/\s*Menanggapi surat/gi, '\n\nMenanggapi surat')
    .replace(/\s*dengan ini kami menyampaikan/gi, '\n\nDengan ini kami menyampaikan')
    .replace(/\s*Hari,\s*tanggal\s*:?/gi, '\n\nHari, tanggal:')
    .replace(/\s*Waktu\s*:?/gi, '\nWaktu:')
    .replace(/\s*Tempat\s*:?/gi, '\nTempat:')
    .replace(/\s*Atas perhatian/gi, '\n\nAtas perhatian')
    .replace(/\s*Hormat kami,?/gi, '\n\nHormat kami,')
    .replace(/\s*Tembusan Yth\.?\s*:?/gi, '\n\nTembusan Yth.:')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  // Preserve manual paragraph separation from plain-text templates.
  return text
    .split(/\n\s*\n/)
    .map(function(block) {
      return '<p style="margin:0 0 12px 0;line-height:1.6">' +
        escapeHtml_(block).replace(/\n/g, '<br>') +
        '</p>';
    })
    .join('');
}

function isExplicitHtmlTemplate_(value) {
  const text = String(value || '').trim();
  if (!text || text.indexOf('<') === -1 || text.indexOf('>') === -1) return false;

  // Only treat as HTML when common tags are present.
  return /<(?:!doctype|html|head|body|title|style|script|p|div|span|br|hr|table|thead|tbody|tr|td|th|ul|ol|li|strong|em|b|i|u|h[1-6]|a|blockquote)\b|<\/[a-z][^>]*>/i.test(text);
}

function createEmailDraftInternal_(documentId, force, user, customRecipientsJson) {
  const sheet = getSheet_('DOCUMENTS');
  const rowNumber = findRowById_(sheet, documentId, 1);
  if (!rowNumber) throw new Error('Dokumen tidak ditemukan: ' + documentId);
  const row = sheet.getRange(rowNumber, 1, 1, DOCUMENT_HEADERS.length).getValues()[0];
  const document = documentRowToDto_(row);
  const request = getRequestDetailInternal_(document.requestId).request;
  if (request.status !== 'READY') {
    throw new Error(
      request.status === 'ARCHIVED'
        ? 'Permohonan sudah selesai dan tidak dapat diproses ulang.'
        : 'Tandai permohonan sebagai Siap Diproses sebelum membuat draft email.'
    );
  }
  const marker = 'DPSP-DRAFT:' + document.id + ':R' + document.revision;

  if (!force && document.emailDraftId) {
    try {
      const existing = GmailApp.getDraft(document.emailDraftId);
      if (existing) {
        return {
          ok: true,
          reused: true,
          requestId: document.requestId,
          documentId: document.id,
          draftId: existing.getId(),
          gmailUrl: 'https://mail.google.com/mail/u/0/#drafts'
        };
      }
    } catch (error) {
      console.warn('Draft ID lama tidak ditemukan: ' + document.emailDraftId);
    }
  }

  if (!force) {
    const recovered = findExistingDraftByMarker_(marker);
    if (recovered) {
      saveDraftState_(sheet, rowNumber, row, recovered.getId());
      return {
        ok: true,
        reused: true,
        recovered: true,
        requestId: document.requestId,
        documentId: document.id,
        draftId: recovered.getId(),
        gmailUrl: 'https://mail.google.com/mail/u/0/#drafts'
      };
    }
  }

  if (document.status !== 'GENERATED' || !document.pdfId || !driveFileExists_(document.pdfId)) {
    throw new Error('Buat dokumen dan PDF sebelum membuat draft email.');
  }

  let customRecipients = null;
  if (customRecipientsJson) {
    try {
      customRecipients = JSON.parse(customRecipientsJson);
    } catch (e) {
      console.error('Failed to parse customRecipientsJson: ' + e.message);
    }
  }
  const preview = buildEmailPreviewForDocument_(document, customRecipients);
  const markerHtml = '<span style="display:none;color:transparent;font-size:0">' +
    escapeHtml_(marker) + '</span>';
  const htmlBody = preview.htmlBody + markerHtml;
  const attachment = DriveApp.getFileById(document.pdfId).getBlob();
  const draft = GmailApp.createDraft(
    preview.to.join(','),
    preview.subject,
    stripHtml_(htmlBody),
    {
      htmlBody: htmlBody,
      cc: preview.cc.join(','),
      bcc: (preview.bcc || []).join(','),
      attachments: [attachment],
      markImportant: true
    }
  );

  saveDraftState_(sheet, rowNumber, row, draft.getId());
  recordDraftArtifact_(document, draft.getId(), user.email, marker);
  updateMasterEmailStatus_(document.requestId);
  return {
    ok: true,
    reused: false,
    requestId: document.requestId,
    documentId: document.id,
    draftId: draft.getId(),
    gmailUrl: 'https://mail.google.com/mail/u/0/#drafts',
    preview: preview
  };
}

function getEmailTemplate_(documentType, speakerStatus) {
  const sheet = getSheet_('EMAIL_TEMPLATE');
  const rows = readDataRows_(sheet, 5);
  let row = rows.find(function(item) {
    return text_(item[0]) === documentType &&
      text_(item[1]) === text_(speakerStatus);
  });
  if (!row) {
    row = rows.find(function(item) {
      return text_(item[0]) === documentType && !text_(item[1]);
    });
  }
  if (!row) throw new Error('Template email tidak ditemukan untuk ' + documentType);
  return { subject: text_(row[2]), body: String(row[3] || '') };
}

function buildEmailPlaceholders_(detail, document, routing) {
  const request = detail.request;
  const employees = detail.employees;
  const employeeJoin = buildEmployeeJoinPlaceholders_(employees);
  const plain = {
    kepadaYth: routing.toRoles.join('\n'),
    jenisSurat: document.type,
    subTipe: document.speakerSubtype || '',
    statusNarasumber: document.speakerStatus || '',
    namaKegiatan: request.activityName,
    namaMitra: request.partnerName,
    narasumber: employeeJoin.narasumber,
    textJoinNomor: employeeJoin.textJoinNomor,
    textJoinNama: employeeJoin.textJoinNama,
    textJoinNikNpm: employeeJoin.textJoinNikNpm,
    textJoinJabatan: employeeJoin.textJoinJabatan,
    textJoinProdi: employeeJoin.textJoinProdi,
    textJoinFakultas: employeeJoin.textJoinFakultas,
    textJoinEmail: employeeJoin.textJoinEmail,
    'Text Join Nomor': employeeJoin.textJoinNomor,
    'Text Join Nama': employeeJoin.textJoinNama,
    'Text Join NIK/NPM': employeeJoin.textJoinNikNpm,
    'Text Join Jabatan': employeeJoin.textJoinJabatan,
    'Text Join Prodi': employeeJoin.textJoinProdi,
    'Text Join Fakultas': employeeJoin.textJoinFakultas,
    'Text Join Email': employeeJoin.textJoinEmail,
    hari: request.day,
    tanggal: request.dateDisplay,
    tempat: request.activityPlace,
    tembusan: routing.ccRoles.join('\n'),
    nomorSurat: document.number,
    nomorSuratMasuk: request.incomingNumber,
    tanggalSuratMasuk: request.incomingDate ? formatIndonesianDate_(request.incomingDate) : '',
    pengirimSuratMasuk: request.incomingSender,
    perihalSuratMasuk: request.incomingSubject,
    alamatMitra: request.partnerAddress,
    waktuKegiatan: request.activityTime
  };
  const output = {};
  Object.keys(plain).forEach(function(key) {
    output[key] = {
      plain: String(plain[key] || ''),
      html: escapeHtml_(plain[key]).replace(/\n/g, '<br>')
    };
  });
  return output;
}

function replaceTemplateTokens_(template, placeholders, html) {
  let output = String(template || '');
  Object.keys(placeholders).forEach(function(key) {
    const value = html ? placeholders[key].html : placeholders[key].plain;
    ['{' + key + '}', '{{' + key + '}}', '<<' + key + '>>'].forEach(function(pattern) {
      output = output.replace(new RegExp(escapeRegex_(pattern), 'g'), value);
    });
  });
  return output;
}

function findExistingDraftByMarker_(marker) {
  const drafts = GmailApp.getDrafts();
  const max = Math.min(drafts.length, 200);
  for (let i = 0; i < max; i++) {
    try {
      const body = drafts[i].getMessage().getBody();
      if (body && body.indexOf(marker) !== -1) return drafts[i];
    } catch (error) {
      console.warn('Gagal membaca draft untuk rekonsiliasi: ' + error.message);
    }
  }
  return null;
}

function saveDraftState_(sheet, rowNumber, row, draftId) {
  row[12] = draftId;
  row[13] = 'DRAFTED';
  row[16] = new Date();
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
}

function recordDraftArtifact_(document, draftId, userEmail, marker) {
  appendDataRow_(getSheet_('FILES'), [
    document.requestId,
    document.templateKey,
    document.revision,
    'EMAIL_DRAFT',
    draftId,
    'https://mail.google.com/mail/u/0/#drafts',
    new Date(),
    userEmail,
    'ACTIVE',
    JSON.stringify({ documentId: document.id, marker: marker })
  ]);
}

function updateMasterEmailStatus_(requestId) {
  const documents = getDocumentsByRequest_(requestId);
  const drafted = documents.length > 0 && documents.every(function(document) {
    return document.emailStatus === 'DRAFTED';
  });
  const sheet = getSheet_('MASTER');
  const rowNumber = findRowById_(sheet, requestId, 1);
  if (!rowNumber) return;
  sheet.getRange(rowNumber, masterColumn_('Email Status'))
    .setValue(drafted ? 'EMAIL DRAFTED' : 'PARTIAL');
}

function getDocumentById_(documentId) {
  const sheet = getSheet_('DOCUMENTS');
  const rowNumber = findRowById_(sheet, documentId, 1);
  if (!rowNumber) throw new Error('Dokumen tidak ditemukan: ' + documentId);
  return documentRowToDto_(
    sheet.getRange(rowNumber, 1, 1, DOCUMENT_HEADERS.length).getValues()[0]
  );
}

function stripHtml_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .trim();
}

function isValidEmail_(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email).trim().toLowerCase());
}

function parseAndValidateEmailString_(str) {
  if (!str) return [];
  const parts = str.split(/[,;\n\s]+/).map(function(p) { return p.trim(); }).filter(Boolean);
  parts.forEach(function(email) {
    if (!isValidEmail_(email)) {
      throw new Error('Format email tidak valid: "' + email + '"');
    }
  });
  return parts;
}

function saveFinalRecipientsToMaster_(requestId, toList, ccList, documentId) {
  try {
    if (documentId) {
      const docSheet = getSheet_('DOCUMENTS');
      const docRowNumber = findRowById_(docSheet, documentId, 1);
      if (docRowNumber) {
        const row = docSheet.getRange(docRowNumber, 1, 1, DOCUMENT_HEADERS.length).getValues()[0];
        row[17] = toList.join(', ');
        row[18] = ccList.join(', ');
        docSheet.getRange(docRowNumber, 1, 1, row.length).setValues([row]);
      }
    }
  } catch (e) {
    console.error('Gagal menyimpan penerima email final: ' + e.message);
  }
}

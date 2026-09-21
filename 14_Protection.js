/**********************************************************************
 * BẢO VỆ SHEET - SV CHỈ ĐƯỢC SỬA ĐÚNG VÙNG CẦN THIẾT
 **********************************************************************/

function applyAllProtections_(ss) {
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.HOME), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.CONFIG), [], true);
  applyRosterProtectionsOnly_(ss, true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.DASHBOARD), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.RULES), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.GENERAL_HISTORY), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.WEEK_HISTORY), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.SYSTEM_LOG), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.SYSTEM), [], true);
  applyRegistrationAccessProtections_(ss, true);

  try { ss.getSheetByName(LAB.SHEETS.SYSTEM).hideSheet(); } catch (err) {}
}

function applyRegistrationAccessProtections_(ss, cleanupLegacyOwners) {
  const phase = getWorkflowPhase_();
  const full = getDocumentProperty_(LAB.PROP.FULL_LOCK) === 'YES';
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const regCount = getRegisterRowCount_(reg);
  const registerOpenRanges = [];

  if (phase === LAB.PHASE_REGISTRATION_OPEN && !full && reg && regCount > 0) {
    registerOpenRanges.push(reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, regCount, 1));
    registerOpenRanges.push(reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.GHICHU_SV, regCount, 1));
  }
  applyManagedSheetBaseProtection_(reg, registerOpenRanges, cleanupLegacyOwners);

  const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const generalCount = getGeneralRowCount_(general);
  const generalOpenRanges = [];
  const generalMode = getWeekMode_() === LAB.MODE_GENERAL_CLEANING;
  if (
    phase === LAB.PHASE_REGISTRATION_OPEN &&
    !full &&
    generalCount > 0 &&
    generalMode
  ) {
    generalOpenRanges.push(general.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, generalCount, 1));
  }
  if (generalMode || cleanupLegacyOwners) {
    applyManagedSheetBaseProtection_(general, generalOpenRanges, cleanupLegacyOwners);
  }

  if (cleanupLegacyOwners) {
    removeManagedRangeProtectionsByPrefix_(reg, LAB.REGISTER_ROW_PREFIX);
    removeManagedRangeProtectionsByPrefix_(general, LAB.GENERAL_ROW_PREFIX);
  }
}

function removeManagedRangeProtectionsByPrefix_(sheet, prefix) {
  if (!sheet) return;
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (protection) {
    const description = protection.getDescription() || '';
    if (description.indexOf(prefix) === 0) {
      try { protection.remove(); } catch (err) {}
    }
  });
}

function applyRosterProtectionsOnly_(ss, cleanupLegacyOwners) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!sh) return;
  const inputRows = sh.getMaxRows() - LAB.FIRST_STUDENT_ROW + 1;
  const openRanges = inputRows > 0
    ? [sh.getRange(LAB.FIRST_STUDENT_ROW, 2, inputRows, 1)]
    : [];
  applyManagedSheetBaseProtection_(sh, openRanges, cleanupLegacyOwners);
  if (cleanupLegacyOwners) {
    removeManagedRangeProtectionsByPrefix_(sh, LAB.ROSTER_ROW_PREFIX);
  }
}

function applyManagedSheetBaseProtection_(sheet, unprotectedRanges, refreshEditors) {
  if (!sheet) return;
  const managed = sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter(function (protection) {
      const description = protection.getDescription() || '';
      return description.indexOf(LAB.PROTECTION_PREFIX) === 0 ||
        description === LAB.LEGACY_LOCK_DESC;
    });
  let protection = managed.shift();
  const created = !protection;
  if (!protection) protection = sheet.protect();
  managed.forEach(function (duplicate) {
    try { duplicate.remove(); } catch (err) {}
  });
  if (created || refreshEditors) {
    configureManagedSheetProtection_(protection, sheet, unprotectedRanges);
  } else {
    protection.setUnprotectedRanges(unprotectedRanges || []);
  }
}

function configureManagedSheetProtection_(protection, sheet, unprotectedRanges) {
  protection.setDescription(LAB.PROTECTION_PREFIX + sheet.getName());
  protection.setWarningOnly(false);

  const admins = getAdminEmails_();
  admins.forEach(function (email) {
    try { protection.addEditor(email); } catch (err) {}
  });

  try {
    protection.getEditors().forEach(function (editor) {
      const email = normalizeEmail_(editor.getEmail());
      if (email && admins.length > 0 && admins.indexOf(email) === -1) {
        protection.removeEditor(editor);
      }
    });
  } catch (err) {}

  try {
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  } catch (err) {}

  protection.setUnprotectedRanges(unprotectedRanges || []);
}

function removeManagedProtectionsFromSheet_(sheet) {
  if (!sheet) return;
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE]
    .forEach(function (type) {
      sheet.getProtections(type).forEach(function (protection) {
        const description = protection.getDescription() || '';
        if (
          description.indexOf(LAB.PROTECTION_PREFIX) === 0 ||
          description.indexOf(LAB.ROSTER_ROW_PREFIX) === 0 ||
          description.indexOf(LAB.REGISTER_ROW_PREFIX) === 0 ||
          description.indexOf(LAB.GENERAL_ROW_PREFIX) === 0 ||
          description === LAB.LEGACY_LOCK_DESC
        ) {
          try { protection.remove(); } catch (err) {}
        }
      });
    });
}

function removeAllManagedProtections_(ss) {
  ss.getSheets().forEach(function (sheet) {
    removeManagedProtectionsFromSheet_(sheet);
  });
}


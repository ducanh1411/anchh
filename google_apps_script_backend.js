/**
 * GOOGLE APPS SCRIPT BACKEND - MATERIAL MASTER ERP
 * Triển khai Chế độ A (Dành cho Server thử nghiệm riêng hoặc Server chung của lớp)
 * Phiên bản Schema: erp-attribute-v1
 */

const SCHEMA_VERSION = "erp-attribute-v1";

function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload || {};

    let response = { ok: false, error: { code: "UNKNOWN_ACTION", message: "Hành động không hợp lệ" } };

    switch (action) {
      case 'ping':
        response = { ok: true, data: { service: "Material Master ERP GAS", version: "1.0-training" } };
        break;

      case 'training.status':
        response = {
          ok: true,
          data: {
            enabled: true,
            mode: "TRAINING_PUBLIC",
            canCreate: true,
            attributeSchemaVersion: SCHEMA_VERSION
          }
        };
        break;

      case 'training.list':
        response = handleListMaterials();
        break;

      case 'training.create':
        response = handleCreateMaterial(payload);
        break;

      default:
        response = { ok: false, error: { code: "DENY_BY_DEFAULT", message: "Yêu cầu bị từ chối." } };
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errorResponse = {
      ok: false,
      error: { code: "SERVER_ERROR", message: err.toString() }
    };
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function handleListMaterials() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MATERIALS");
  if (!sheet) {
    return { ok: true, data: { materials: [], count: 0 } };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { ok: true, data: { materials: [], count: 0 } };
  }

  const headers = data[0];
  const materials = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    materials.push({
      id: String(row[0]),
      code: String(row[1]),
      name: String(row[2]),
      categoryId: String(row[3]),
      unit: String(row[4]),
      materialType: String(row[5]),
      grade: String(row[6]),
      standard: String(row[7]),
      dimensions: String(row[8]),
      notes: String(row[10]),
      createdAt: row[14] ? new Date(row[14]).toISOString() : new Date().toISOString(),
      createdBy: String(row[16]),
      attributesJson: String(row[20] || '{}')
    });
  }

  return { ok: true, data: { materials: materials.reverse(), count: materials.length } };
}

function handleCreateMaterial(payload) {
  const lock = LockService.getScriptLock();
  try {
    // Chờ tối đa 10 giây để lấy khóa giao dịch độc quyền
    if (!lock.waitLock(10000)) {
      return { ok: false, error: { code: "LOCK_TIMEOUT", message: "Hệ thống bận, vui lòng thử lại." } };
    }

    const mat = payload.material;
    const actorName = payload.actorName || "Học viên";
    const requestId = payload.requestId;

    if (!mat || !mat.categoryId || !mat.subtypeId || !mat.attributes) {
      return { ok: false, error: { code: "INVALID_INPUT", message: "Thiếu thông tin thuộc tính bắt buộc." } };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let matSheet = ss.getSheetByName("MATERIALS");
    let rangeSheet = ss.getSheetByName("NUMBER_RANGES");

    if (!matSheet || !rangeSheet) {
      setupSheets(ss);
      matSheet = ss.getSheetByName("MATERIALS");
      rangeSheet = ss.getSheetByName("NUMBER_RANGES");
    }

    // 1. Kiểm tra Idempotency requestId
    const matData = matSheet.getDataRange().getValues();
    for (let i = 1; i < matData.length; i++) {
      if (matData[i][19] === requestId) { // Cột T: requestId
        return {
          ok: true,
          data: {
            material: {
              id: String(matData[i][0]),
              code: String(matData[i][1]),
              name: String(matData[i][2]),
              categoryId: String(matData[i][3])
            },
            idempotentReplay: true
          }
        };
      }
    }

    // 2. Server-side Duplicate Identity Signature Check
    const newSignature = buildIdentitySignature(mat.categoryId, mat.subtypeId, mat.attributes);
    for (let i = 1; i < matData.length; i++) {
      const existingCat = matData[i][3];
      const existingAttrJson = matData[i][20];
      if (existingCat === mat.categoryId && existingAttrJson) {
        try {
          const existingAttrs = JSON.parse(existingAttrJson);
          const existingSubtype = existingAttrs._subtypeId || "";
          if (existingSubtype === mat.subtypeId) {
            const existingSig = buildIdentitySignature(existingCat, existingSubtype, existingAttrs);
            if (existingSig === newSignature) {
              return {
                ok: false,
                error: {
                  code: "DUPLICATE_MATERIAL",
                  message: `Vật tư đã tồn tại trong danh mục với Mã ${matData[i][1]}`
                }
              };
            }
          }
        } catch (e) {}
      }
    }

    // 3. Cấp mã mới từ NUMBER_RANGES
    const rangeData = rangeSheet.getDataRange().getValues();
    let lastNum = 0;
    let rangeRowIndex = 2;

    for (let r = 1; r < rangeData.length; r++) {
      if (rangeData[r][0] === 'MATERIAL_AUTO') {
        lastNum = parseInt(rangeData[r][3]) || 0;
        rangeRowIndex = r + 1;
        break;
      }
    }

    const nextNum = lastNum + 1;
    const newCode = "A" + String(nextNum).padStart(6, '0');
    const generatedName = generateServerName(mat.categoryId, mat.subtypeId, mat.attributes);

    // Ghi lưu số tiếp theo
    rangeSheet.getRange(rangeRowIndex, 4).setValue(nextNum);
    rangeSheet.getRange(rangeRowIndex, 5).setValue("A" + String(nextNum + 1).padStart(6, '0'));
    rangeSheet.getRange(rangeRowIndex, 7).setValue(new Date().toISOString());

    // Save attributes enriched with subtypeId
    const finalAttrs = Object.assign({}, mat.attributes, { _subtypeId: mat.subtypeId });

    // 4. Ghi bản ghi vào Tab MATERIALS
    const rowId = String(Date.now());
    const nowIso = new Date().toISOString();

    matSheet.appendRow([
      rowId,                   // A: id
      newCode,                 // B: code
      generatedName,           // C: name
      mat.categoryId,          // D: categoryId
      mat.unit || 'Cái',       // E: unit
      mat.materialType || 'ROH',// F: materialType
      mat.attributes.grade || '', // G: grade
      mat.attributes.standard || '', // H: standard
      "",                      // I: dimensions
      "",                      // J: techSpec
      mat.notes || '',         // K: notes
      "ACTIVE",                // L: status
      "SYSTEM_GENERATED",      // M: codeOrigin
      "TRAINING",              // N: origin
      nowIso,                  // O: createdAt
      nowIso,                  // P: updatedAt
      actorName,               // Q: createdBy
      actorName,               // R: updatedBy
      1,                       // S: revision
      requestId,               // T: requestId
      JSON.stringify(finalAttrs)// U: attributesJson
    ]);

    return {
      ok: true,
      data: {
        material: {
          id: rowId,
          code: newCode,
          name: generatedName,
          categoryId: mat.categoryId,
          subtypeId: mat.subtypeId,
          unit: mat.unit || 'Cái',
          materialType: mat.materialType || 'ROH',
          createdAt: nowIso,
          createdBy: actorName,
          attributesJson: JSON.stringify(finalAttrs)
        },
        idempotentReplay: false
      }
    };

  } finally {
    lock.releaseLock();
  }
}

function buildIdentitySignature(catId, subId, attrs) {
  const cleanObj = { categoryId: catId, subtypeId: subId };
  Object.keys(attrs).sort().forEach(k => {
    if (k !== '_subtypeId') {
      cleanObj[k] = String(attrs[k]).trim().toUpperCase();
    }
  });
  return JSON.stringify(cleanObj);
}

function generateServerName(catId, subId, attrs) {
  const getV = (k) => attrs[k] ? String(attrs[k]).trim() : '';
  if (subId === 'STEEL_PLATE') {
    return `Thép tấm ${getV('grade')}${getV('standard') ? ' (' + getV('standard') + ')' : ''} · ${getV('thickness')}x${getV('width')}x${getV('length')} mm`;
  }
  if (subId === 'HEX_BOLT') {
    return `Bu lông lục giác ngoài ${getV('standard')} · ${getV('strength')} · M${getV('diameter')}x${getV('length')} · ${getV('finish')}`;
  }
  return `Vật tư ${catId} · ${getV('grade') || getV('compound') || ''}`;
}

function setupSheets(ss) {
  let matSheet = ss.getSheetByName("MATERIALS");
  if (!matSheet) {
    matSheet = ss.insertSheet("MATERIALS");
    matSheet.appendRow([
      "id", "code", "name", "categoryId", "unit", "materialType", "grade", "standard",
      "dimensions", "technicalSpecification", "notes", "status", "codeOrigin", "origin",
      "createdAt", "updatedAt", "createdBy", "updatedBy", "revision", "requestId", "attributesJson"
    ]);
  }

  let rangeSheet = ss.getSheetByName("NUMBER_RANGES");
  if (!rangeSheet) {
    rangeSheet = ss.insertSheet("NUMBER_RANGES");
    rangeSheet.appendRow(["series", "prefix", "digits", "lastNumber", "nextCode", "enabled", "updatedAt"]);
    rangeSheet.appendRow(["MATERIAL_AUTO", "A", 6, 0, "A000001", true, new Date().toISOString()]);
  }
}
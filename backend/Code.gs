// ===== GLOBAL CONFIG =====
const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

const PARTNER_SHEET_NAMES = Object.freeze({
  partners: "partners",
  partnerLinks: "partner_links",
  partnerCatalogItems: "partner_catalog_items",
  partnerRequests: "partner_requests",
  partnerRequestItems: "partner_request_items"
});

const PARTNER_SHEET_SCHEMAS = Object.freeze({
  partners: [
    "partnerId",
    "partnerName",
    "contactName",
    "email",
    "phone",
    "company",
    "status",
    "note",
    "createdAt",
    "createdBy",
    "updatedAt",
    "updatedBy"
  ],
  partner_links: [
    "linkId",
    "partnerId",
    "tokenHash",
    "status",
    "expiresAt",
    "label",
    "createdAt",
    "createdBy",
    "disabledAt",
    "disabledBy",
    "lastAccessedAt",
    "accessCount",
    "renewedFromLinkId"
  ],
  partner_catalog_items: [
    "linkId",
    "productId",
    "visible",
    "sortOrder",
    "note",
    "createdAt",
    "createdBy"
  ],
  partner_requests: [
    "requestId",
    "linkId",
    "partnerId",
    "partnerNameSnapshot",
    "contactName",
    "contactEmail",
    "contactPhone",
    "message",
    "status",
    "itemCount",
    "estimatedTotal",
    "submittedAt",
    "sourceIpHash",
    "userAgent",
    "reviewedAt",
    "reviewedBy",
    "adminNote"
  ],
  partner_request_items: [
    "requestId",
    "productId",
    "nameSnapshot",
    "priceSnapshot",
    "qty",
    "statusSnapshot",
    "imageSnapshot",
    "partnerNote",
    "sortOrder"
  ]
});

const PARTNER_REQUEST_MAX_ITEMS = 100;
const PARTNER_REQUEST_MAX_CONTACT_LENGTH = 120;
const PARTNER_REQUEST_MAX_MESSAGE_LENGTH = 1000;
const PARTNER_REQUEST_MAX_NOTE_LENGTH = 500;
const PARTNER_REQUEST_ALLOWED_STATUSES = Object.freeze([
  "new",
  "reviewed",
  "closed",
  "cancelled"
]);
const PARTNER_LINK_ALLOWED_STATUSES = Object.freeze([
  "active",
  "disabled"
]);
const ORDER_OPTIONAL_METADATA_HEADERS = Object.freeze([
  "source",
  "partnerRequestId",
  "partnerId",
  "partnerNameSnapshot"
]);
const PRODUCT_OPTIONAL_VARIANT_HEADERS = Object.freeze([
  "parentProductId",
  "variantType",
  "variantValue",
  "displayName",
  "variantSortOrder"
]);

function ensurePartnerCatalogSheets() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    return {
      success: true,
      sheets: ensurePartnerCatalogSheets_(ss)
    };
  } finally {
    lock.releaseLock();
  }
}

function ensurePartnerCatalogSheets_(ss) {
  const results = [];

  Object.keys(PARTNER_SHEET_SCHEMAS).forEach(name => {
    results.push(
      ensureSheetWithHeaders_(
        ss,
        name,
        PARTNER_SHEET_SCHEMAS[name]
      )
    );
  });

  return results;
}

function ensureSheetWithHeaders_(ss, sheetName, expectedHeaders) {
  if (!ss) {
    throw new Error("Spreadsheet is required");
  }

  if (!sheetName || !Array.isArray(expectedHeaders) || !expectedHeaders.length) {
    throw new Error("Invalid sheet schema");
  }

  let sheet = ss.getSheetByName(sheetName);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    created = true;
  }

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, expectedHeaders.length)
      .setValues([expectedHeaders]);
    sheet.setFrozenRows(1);
    return {
      name: sheetName,
      created,
      initialized: true
    };
  }

  const actualHeaders = sheet
    .getRange(1, 1, 1, expectedHeaders.length)
    .getValues()[0]
    .map(header => String(header || "").trim());

  const matches = expectedHeaders.every((header, index) =>
    actualHeaders[index] === header
  );

  if (!matches) {
    throw new Error(sheetName + " schema mismatch");
  }

  return {
    name: sheetName,
    created,
    initialized: false
  };
}

function ensureOrdersExtendedHeaders_(ss) {
  if (!ss) {
    throw new Error("Spreadsheet is required");
  }

  const sheet = ss.getSheetByName("Orders");
  if (!sheet) {
    throw new Error("Orders sheet not found");
  }

  if (sheet.getLastRow() === 0) {
    throw new Error("Orders sheet has no header row");
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(header => String(header || "").trim());
  const missingHeaders = ORDER_OPTIONAL_METADATA_HEADERS.filter(
    header => headers.indexOf(header) === -1
  );

  if (missingHeaders.length) {
    sheet
      .getRange(1, lastColumn + 1, 1, missingHeaders.length)
      .setValues([missingHeaders]);
  }

  return {
    appended: missingHeaders,
    headers: headers.concat(missingHeaders)
  };
}

function ensureProductsVariantHeaders_(ss) {
  if (!ss) {
    throw new Error("Spreadsheet is required");
  }

  const sheet = ss.getSheetByName("Products");
  if (!sheet) {
    throw new Error("Products sheet not found");
  }

  if (sheet.getLastRow() === 0) {
    throw new Error("Products sheet has no header row");
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(header => String(header || "").trim());
  const missingHeaders = PRODUCT_OPTIONAL_VARIANT_HEADERS.filter(
    header => headers.indexOf(header) === -1
  );

  if (missingHeaders.length) {
    sheet
      .getRange(1, lastColumn + 1, 1, missingHeaders.length)
      .setValues([missingHeaders]);
  }

  return {
    appended: missingHeaders,
    headers: headers.concat(missingHeaders)
  };
}

function getProductVariantColumnMap_(headers) {
  const sourceHeaders = Array.isArray(headers) ? headers : [];
  const col = {};
  PRODUCT_OPTIONAL_VARIANT_HEADERS.forEach(header => {
    col[header] = sourceHeaders.indexOf(header);
  });
  return col;
}

function getProductVariantMetadata_(row, col) {
  const columnMap = col || {};
  const sortValue = getProductVariantValue_(row, columnMap.variantSortOrder);
  const sortNumber = Number(sortValue);

  return {
    parentProductId: getProductVariantValue_(row, columnMap.parentProductId),
    variantType: getProductVariantValue_(row, columnMap.variantType),
    variantValue: getProductVariantValue_(row, columnMap.variantValue),
    displayName: getProductVariantValue_(row, columnMap.displayName),
    variantSortOrder:
      sortValue !== "" && Number.isFinite(sortNumber)
        ? sortNumber
        : sortValue
  };
}

function getProductVariantValue_(row, index) {
  if (!row || index === undefined || index < 0) {
    return "";
  }
  const value = row[index];
  if (value === undefined || value === null) {
    return "";
  }
  return String(value).trim();
}

function parseProductVariantMetadata_(params) {
  const source = params || {};
  const rawSortOrder = String(source.variantSortOrder ?? "").trim();
  let variantSortOrder = "";

  if (rawSortOrder !== "") {
    const sortOrder = Number(rawSortOrder);
    if (!Number.isInteger(sortOrder)) {
      throw new Error("Variant sort order must be an integer");
    }
    variantSortOrder = sortOrder;
  }

  return {
    parentProductId: String(source.parentProductId || "").trim(),
    variantType: String(source.variantType || "").trim(),
    variantValue: String(source.variantValue || "").trim(),
    displayName: String(source.displayName || "").trim(),
    variantSortOrder
  };
}

function writeProductVariantMetadata_(sheet, rowNumber, metadata, headers) {
  if (!sheet || !rowNumber || !metadata) return;

  const sourceHeaders = Array.isArray(headers) && headers.length
    ? headers
    : ensureProductsVariantHeaders_(getSS()).headers;

  PRODUCT_OPTIONAL_VARIANT_HEADERS.forEach(header => {
    const columnIndex = sourceHeaders.indexOf(header);
    if (columnIndex === -1) return;
    const value = metadata[header] === undefined || metadata[header] === null
      ? ""
      : metadata[header];
    sheet.getRange(rowNumber, columnIndex + 1).setValue(value);
  });
}

function findOrderByPartnerRequestId_(orderSheet, orderCol, requestId) {
  if (
    !orderSheet ||
    !orderCol ||
    orderCol.partnerRequestId === undefined
  ) {
    return null;
  }

  const targetRequestId = String(requestId || "").trim();
  if (!targetRequestId || orderSheet.getLastRow() < 2) {
    return null;
  }

  const values = orderSheet
    .getRange(
      2,
      orderCol.partnerRequestId + 1,
      orderSheet.getLastRow() - 1,
      1
    )
    .getValues();
  const matchIndex = values.findIndex(row =>
    String(row[0] || "").trim() === targetRequestId
  );

  if (matchIndex === -1) {
    return null;
  }

  return {
    rowNumber: matchIndex + 2
  };
}

function createPartnerLink(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const partnerId = String(params.partnerId || "").trim();
  const label = String(params.label || "").trim();
  const expiresAt = parseFuturePartnerDate_(params.expiresAt);
  const productIds = parsePartnerProductIds_(params.productIds);

  if (!partnerId) {
    throw new Error("Partner ID is required");
  }

  if (!productIds.length) {
    throw new Error("At least one product is required");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    ensurePartnerCatalogSheets_(ss);

    const partnerData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partners
    );
    const linkData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partnerLinks
    );
    const catalogSheet = ss.getSheetByName(
      PARTNER_SHEET_NAMES.partnerCatalogItems
    );
    const productSheet = ss.getSheetByName("Products");

    if (!productSheet) {
      throw new Error("Products sheet not found");
    }

    const partner = findPartnerById_(partnerData, partnerId);
    if (!partner) {
      throw new Error("Partner not found or disabled");
    }

    const productMap = getProductRowMap_(productSheet);
    productIds.forEach(productId => {
      if (!productMap.has(productId)) {
        throw new Error("Product not found: " + productId);
      }
    });

    const linkId = "PL-" + Utilities.getUuid();
    const rawToken = createPartnerRawToken_();
    const tokenHash = hashPartnerToken_(rawToken);

    if (findPartnerLinkByTokenHash_(linkData, tokenHash)) {
      throw new Error("Partner token collision");
    }

    const createdAt = new Date();
    const createdProductIds = [];
    let linkCreated = false;

    try {
      linkData.sheet.appendRow([
        linkId,
        partnerId,
        tokenHash,
        "active",
        expiresAt,
        label,
        createdAt,
        auth.username,
        "",
        "",
        "",
        0,
        ""
      ]);
      linkCreated = true;

      productIds.forEach((productId, index) => {
        catalogSheet.appendRow([
          linkId,
          productId,
          true,
          index + 1,
          "",
          createdAt,
          auth.username
        ]);
        createdProductIds.push(productId);
      });
    } catch (err) {
      rollbackPartnerLinkCreate_(
        linkData.sheet,
        catalogSheet,
        linkId,
        linkCreated,
        createdProductIds.length
      );
      throw err;
    }

    return {
      success: true,
      data: {
        linkId,
        partnerId,
        token: rawToken,
        url: getPartnerCatalogUrl_(rawToken),
        expiresAt
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function getPartnerCatalog(partnerToken) {
  const rawToken = String(partnerToken || "").trim();
  if (!rawToken) {
    throw new Error("Missing partner token");
  }

  const ss = getSS();
  const context = getPartnerCatalogContext_(ss, rawToken);
  const products = context.selectedProductIds
    .map(productId => context.productMap.get(productId))
    .filter(row => row && isProductActive_(row))
    .map(row => toSafePartnerProduct_(row, context.productVariantCol));

  return {
    success: true,
    data: {
      partner: {
        partnerId: context.partner.partnerId,
        partnerName: context.partner.partnerName
      },
      link: {
        linkId: context.link.linkId,
        expiresAt: context.link.expiresAt,
        label: context.link.label || ""
      },
      products
    }
  };
}

function submitPartnerRequest(params) {
  const rawToken = String(params.partnerToken || "").trim();
  if (!rawToken) {
    throw new Error("Missing partner token");
  }

  const requestItems = parsePartnerRequestItems_(params.items);
  const contactName = sanitizePartnerText_(
    params.contactName,
    PARTNER_REQUEST_MAX_CONTACT_LENGTH
  );
  const contactEmail = sanitizePartnerText_(
    params.contactEmail,
    PARTNER_REQUEST_MAX_CONTACT_LENGTH
  );
  const contactPhone = sanitizePartnerText_(
    params.contactPhone,
    PARTNER_REQUEST_MAX_CONTACT_LENGTH
  );
  const message = sanitizePartnerText_(
    params.message,
    PARTNER_REQUEST_MAX_MESSAGE_LENGTH
  );

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    const context = getPartnerCatalogContext_(ss, rawToken);
    const requestSheet = ss.getSheetByName(
      PARTNER_SHEET_NAMES.partnerRequests
    );
    const itemSheet = ss.getSheetByName(
      PARTNER_SHEET_NAMES.partnerRequestItems
    );

    if (!requestSheet) {
      throw new Error("partner_requests sheet not found");
    }
    if (!itemSheet) {
      throw new Error("partner_request_items sheet not found");
    }

    const visibleProductIds = new Set(context.selectedProductIds);
    const seen = {};
    const cleanItems = requestItems.map((item, index) => {
      const productId = String(item.productId || "").trim().toUpperCase();
      const qty = item.qty;

      if (!visibleProductIds.has(productId)) {
        throw new Error("Product not available for this partner link: " + productId);
      }
      if (seen[productId]) {
        throw new Error("Duplicate productId in request: " + productId);
      }
      seen[productId] = true;

      const productRow = context.productMap.get(productId);
      if (!productRow || !isProductActive_(productRow)) {
        throw new Error("Product not available for this partner link: " + productId);
      }

      const safeProduct = toSafePartnerProduct_(productRow);
      return {
        productId,
        nameSnapshot: safeProduct.name,
        priceSnapshot: safeProduct.price,
        qty,
        statusSnapshot: safeProduct.status,
        imageSnapshot: safeProduct.image,
        partnerNote: sanitizePartnerText_(
          item.partnerNote,
          PARTNER_REQUEST_MAX_NOTE_LENGTH
        ),
        sortOrder: index + 1
      };
    });

    const estimatedTotal = cleanItems.reduce(
      (sum, item) => sum + item.priceSnapshot * item.qty,
      0
    );
    const requestId = "PR-" + Utilities.getUuid();
    const submittedAt = new Date();
    let headerCreated = false;
    let itemRowsCreated = 0;

    try {
      requestSheet.appendRow([
        requestId,
        context.link.linkId,
        context.partner.partnerId,
        context.partner.partnerName,
        contactName,
        contactEmail,
        contactPhone,
        message,
        "new",
        cleanItems.length,
        estimatedTotal,
        submittedAt,
        "",
        "",
        "",
        "",
        ""
      ]);
      headerCreated = true;

      cleanItems.forEach(item => {
        itemSheet.appendRow([
          requestId,
          item.productId,
          item.nameSnapshot,
          item.priceSnapshot,
          item.qty,
          item.statusSnapshot,
          item.imageSnapshot,
          item.partnerNote,
          item.sortOrder
        ]);
        itemRowsCreated++;
      });
    } catch (err) {
      rollbackPartnerRequestCreate_(
        requestSheet,
        itemSheet,
        requestId,
        headerCreated,
        itemRowsCreated
      );
      throw err;
    }

    return {
      success: true,
      data: {
        requestId,
        submittedAt,
        itemCount: cleanItems.length,
        estimatedTotal
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function listPartnerRequests(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const statusFilter = String(params.status || "").trim().toLowerCase();
  const partnerIdFilter = String(params.partnerId || "").trim();
  const searchText = String(params.q || "").trim().toLowerCase();
  const limit = clampPartnerRequestLimit_(params.limit);
  const offset = clampPartnerRequestOffset_(params.offset);

  const ss = getSS();
  const requestData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerRequests
  );

  const requests = requestData.rows
    .map(row => mapPartnerRequestRow_(requestData, row))
    .filter(request => request.requestId)
    .filter(request => {
      if (statusFilter && statusFilter !== "all") {
        if (String(request.status || "").toLowerCase() !== statusFilter) {
          return false;
        }
      }

      if (partnerIdFilter) {
        if (String(request.partnerId || "").trim() !== partnerIdFilter) {
          return false;
        }
      }

      if (searchText) {
        const haystack = [
          request.requestId,
          request.partnerId,
          request.partnerNameSnapshot,
          request.contactName,
          request.contactEmail,
          request.contactPhone,
          request.message
        ].join(" ").toLowerCase();
        if (haystack.indexOf(searchText) === -1) {
          return false;
        }
      }

      return true;
    })
    .sort((left, right) =>
      getPartnerRequestDateTime_(right.submittedAt) -
      getPartnerRequestDateTime_(left.submittedAt)
    );

  return {
    success: true,
    data: {
      requests: requests.slice(offset, offset + limit),
      total: requests.length,
      limit,
      offset
    }
  };
}

function getPartnerRequestDetail(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const requestId = String(params.requestId || "").trim();
  if (!requestId) {
    throw new Error("Request ID is required");
  }

  const ss = getSS();
  const requestData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerRequests
  );
  const itemData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerRequestItems
  );

  const requestRow = requestData.rows.find(row =>
    String(row[requestData.col.requestId] || "").trim() === requestId
  );

  if (!requestRow) {
    throw new Error("Partner request not found");
  }

  const request = mapPartnerRequestRow_(requestData, requestRow);
  const items = itemData.rows
    .filter(row =>
      String(row[itemData.col.requestId] || "").trim() === requestId
    )
    .map(row => mapPartnerRequestItemRow_(itemData, row))
    .sort((left, right) =>
      (Number(left.sortOrder) || 0) - (Number(right.sortOrder) || 0)
    )
    .map(item => {
      const copy = Object.assign({}, item);
      delete copy.sortOrder;
      return copy;
    });

  return {
    success: true,
    data: {
      request,
      items
    }
  };
}

function updatePartnerRequestStatus(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const requestId = String(params.requestId || "").trim();
  const status = String(params.status || "").trim().toLowerCase();

  if (!requestId) {
    throw new Error("Request ID is required");
  }

  if (!status) {
    throw new Error("Status is required");
  }

  if (PARTNER_REQUEST_ALLOWED_STATUSES.indexOf(status) === -1) {
    throw new Error("Invalid partner request status");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    const requestData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partnerRequests
    );

    const rowIndex = requestData.rows.findIndex(row =>
      String(row[requestData.col.requestId] || "").trim() === requestId
    );

    if (rowIndex === -1) {
      throw new Error("Partner request not found");
    }

    requestData.sheet
      .getRange(rowIndex + 2, requestData.col.status + 1)
      .setValue(status);

    const updatedRow = requestData.rows[rowIndex].slice();
    updatedRow[requestData.col.status] = status;

    return {
      success: true,
      data: {
        request: mapPartnerRequestRow_(requestData, updatedRow)
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function convertPartnerRequestToOrder(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const requestId = String(params.requestId || "").trim();
  const poNumber = String(params.poNumber || "").trim();

  if (!requestId) {
    throw new Error("Request ID is required");
  }

  if (!poNumber) {
    throw new Error("PO number is required");
  }

  if (poNumber.length > CREATE_ORDER_MAX_PO_LENGTH) {
    throw new Error(
      `PO number cannot exceed ${CREATE_ORDER_MAX_PO_LENGTH} characters`
    );
  }

  if (/^[=+\-@]/.test(poNumber)) {
    throw new Error("PO number contains an unsafe leading character");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    const requestData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partnerRequests
    );
    const itemData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partnerRequestItems
    );
    const productSheet = ss.getSheetByName("Products");
    const orderSheet = ss.getSheetByName("Orders");

    if (!productSheet) {
      throw new Error("Products sheet not found");
    }

    if (!orderSheet) {
      throw new Error("Orders sheet not found");
    }

    const orderHeaders = ensureOrdersExtendedHeaders_(ss).headers;
    const orderCol = {};
    orderHeaders.forEach((header, index) => {
      orderCol[header] = index;
    });

    const requestRow = requestData.rows.find(row =>
      String(row[requestData.col.requestId] || "").trim() === requestId
    );

    if (!requestRow) {
      throw new Error("Partner request not found");
    }

    const request = mapPartnerRequestRow_(requestData, requestRow);
    const requestStatus = String(request.status || "").trim().toLowerCase();

    if (requestStatus === "cancelled") {
      throw new Error("Cancelled partner request cannot be converted");
    }

    const duplicateOrder = findOrderByPartnerRequestId_(
      orderSheet,
      orderCol,
      requestId
    );

    if (duplicateOrder) {
      throw new Error("Partner request has already been converted");
    }

    const requestItems = itemData.rows
      .filter(row =>
        String(row[itemData.col.requestId] || "").trim() === requestId
      )
      .map(row => mapPartnerRequestItemRow_(itemData, row))
      .sort((left, right) =>
        (Number(left.sortOrder) || 0) - (Number(right.sortOrder) || 0)
      );

    if (!requestItems.length) {
      throw new Error("Partner request has no items");
    }

    const productMap = getProductRowMap_(productSheet);
    const cleanItems = [];
    let total = 0;

    requestItems.forEach((item, index) => {
      const productId = String(item.productId || "").trim().toUpperCase();
      const qty = Number(item.qty);

      if (!productId || !Number.isInteger(qty) || qty <= 0) {
        throw new Error("Invalid partner request item at index " + index);
      }

      const productRow = productMap.get(productId);
      if (!productRow) {
        throw new Error("Product not found: " + productId);
      }

      if (!isProductActive_(productRow)) {
        throw new Error("Product is not active: " + productId);
      }

      const name = String(productRow[1] || "").trim();
      const price = Number(productRow[2]);
      const currentStock = Number(productRow[3]);
      const costPrice = Number(productRow[12]);

      if (
        !name ||
        !Number.isFinite(price) ||
        price < 0 ||
        !Number.isFinite(costPrice) ||
        costPrice < 0
      ) {
        throw new Error("Invalid product data: " + productId);
      }

      if (!Number.isFinite(currentStock) || currentStock < qty) {
        throw new Error(
          `Stock not enough for ${productId} (remain ${currentStock})`
        );
      }

      cleanItems.push({
        productId,
        name,
        qty,
        price,
        costPrice
      });

      total += qty * price;
    });

    const orderId = "ORD-" + Utilities.getUuid();
    const status = "PENDING";
    const createdAt = new Date();
    const row = new Array(orderHeaders.length).fill("");

    row[orderCol.orderId] = orderId;
    row[orderCol.items] = JSON.stringify(cleanItems);
    row[orderCol.total] = total;
    row[orderCol.status] = status;
    row[orderCol.createdAt] = createdAt;
    row[orderCol.approvedAt] = "";
    row[orderCol.approvedBy] = "";
    row[orderCol.poNumber] = poNumber;
    row[orderCol.source] = "partner_request";
    row[orderCol.partnerRequestId] = requestId;
    row[orderCol.partnerId] = request.partnerId;
    row[orderCol.partnerNameSnapshot] = request.partnerNameSnapshot;

    orderSheet.appendRow(row);

    return {
      success: true,
      data: {
        orderId,
        status,
        partnerRequestId: requestId,
        partnerId: request.partnerId,
        partnerNameSnapshot: request.partnerNameSnapshot,
        poNumber,
        total,
        itemCount: cleanItems.length
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function listPartnerLinks(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const partnerIdFilter = String(params.partnerId || "").trim();
  const statusFilter = String(params.status || "").trim().toLowerCase();
  const searchText = String(params.q || "").trim().toLowerCase();
  const limit = clampPartnerRequestLimit_(params.limit);
  const offset = clampPartnerRequestOffset_(params.offset);

  const ss = getSS();
  const linkData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerLinks
  );
  const partnerData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partners
  );
  const catalogData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerCatalogItems
  );

  const partnerNameById = getPartnerNameByIdMap_(partnerData);
  const itemCountByLinkId = getPartnerCatalogItemCountMap_(catalogData);
  const links = linkData.rows
    .map(row =>
      mapPartnerLinkRow_(
        linkData,
        row,
        partnerNameById,
        itemCountByLinkId
      )
    )
    .filter(link => link.linkId)
    .filter(link => {
      if (partnerIdFilter && link.partnerId !== partnerIdFilter) {
        return false;
      }

      if (statusFilter && statusFilter !== "all") {
        if (String(link.status || "").toLowerCase() !== statusFilter) {
          return false;
        }
      }

      if (searchText) {
        const haystack = [
          link.linkId,
          link.partnerId,
          link.partnerNameSnapshot,
          link.label,
          link.status,
          link.createdBy
        ].join(" ").toLowerCase();
        if (haystack.indexOf(searchText) === -1) {
          return false;
        }
      }

      return true;
    })
    .sort((left, right) =>
      getPartnerRequestDateTime_(right.createdAt) -
      getPartnerRequestDateTime_(left.createdAt)
    );

  return {
    success: true,
    data: {
      links: links.slice(offset, offset + limit),
      total: links.length,
      limit,
      offset
    }
  };
}

function getPartnerLinkDetail(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const linkId = String(params.linkId || "").trim();
  if (!linkId) {
    throw new Error("Partner link ID is required");
  }

  const ss = getSS();
  const linkData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerLinks
  );
  const partnerData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partners
  );
  const catalogData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerCatalogItems
  );
  const productSheet = ss.getSheetByName("Products");

  if (!productSheet) {
    throw new Error("Products sheet not found");
  }

  const linkRow = linkData.rows.find(row =>
    String(row[linkData.col.linkId] || "").trim() === linkId
  );

  if (!linkRow) {
    throw new Error("Partner link not found");
  }

  const partnerNameById = getPartnerNameByIdMap_(partnerData);
  const itemCountByLinkId = getPartnerCatalogItemCountMap_(catalogData);
  const productMap = getProductRowMap_(productSheet);
  const link = mapPartnerLinkRow_(
    linkData,
    linkRow,
    partnerNameById,
    itemCountByLinkId
  );
  const items = catalogData.rows
    .filter(row =>
      String(row[catalogData.col.linkId] || "").trim() === linkId &&
      parsePartnerBoolean_(row[catalogData.col.visible])
    )
    .sort((left, right) =>
      (Number(left[catalogData.col.sortOrder]) || 0) -
      (Number(right[catalogData.col.sortOrder]) || 0)
    )
    .map(row => mapPartnerCatalogItemRow_(catalogData, row, productMap));

  return {
    success: true,
    data: {
      link,
      items
    }
  };
}

function updatePartnerLinkStatus(params, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const linkId = String(params.linkId || "").trim();
  const status = String(params.status || "").trim().toLowerCase();

  if (!linkId) {
    throw new Error("Partner link ID is required");
  }

  if (!status) {
    throw new Error("Status is required");
  }

  if (PARTNER_LINK_ALLOWED_STATUSES.indexOf(status) === -1) {
    throw new Error("Invalid partner link status");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    const linkData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partnerLinks
    );
    const partnerData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partners
    );
    const catalogData = getSheetDataBySchema_(
      ss,
      PARTNER_SHEET_NAMES.partnerCatalogItems
    );

    const rowIndex = linkData.rows.findIndex(row =>
      String(row[linkData.col.linkId] || "").trim() === linkId
    );

    if (rowIndex === -1) {
      throw new Error("Partner link not found");
    }

    linkData.sheet
      .getRange(rowIndex + 2, linkData.col.status + 1)
      .setValue(status);

    const updatedRow = linkData.rows[rowIndex].slice();
    updatedRow[linkData.col.status] = status;

    return {
      success: true,
      data: {
        link: mapPartnerLinkRow_(
          linkData,
          updatedRow,
          getPartnerNameByIdMap_(partnerData),
          getPartnerCatalogItemCountMap_(catalogData)
        )
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function mapPartnerLinkRow_(
  linkData,
  row,
  partnerNameById,
  itemCountByLinkId
) {
  const linkId = String(row[linkData.col.linkId] || "").trim();
  const partnerId = String(row[linkData.col.partnerId] || "").trim();
  const link = {
    linkId,
    partnerId,
    partnerNameSnapshot: String(partnerNameById[partnerId] || "").trim(),
    label: String(row[linkData.col.label] || "").trim(),
    status: String(row[linkData.col.status] || "").trim(),
    expiresAt: row[linkData.col.expiresAt] || "",
    createdAt: row[linkData.col.createdAt] || "",
    createdBy: String(row[linkData.col.createdBy] || "").trim(),
    itemCount: Number(itemCountByLinkId[linkId]) || 0
  };
  link.isExpired = isPartnerLinkExpired_(link);
  return link;
}

function mapPartnerCatalogItemRow_(catalogData, row, productMap) {
  const productId = String(row[catalogData.col.productId] || "")
    .trim()
    .toUpperCase();
  const productRow = productMap.get(productId);
  const product = productRow
    ? toSafePartnerProduct_(productRow)
    : null;

  return {
    linkId: String(row[catalogData.col.linkId] || "").trim(),
    productId,
    nameSnapshot: product ? product.name : "",
    priceSnapshot: product ? product.price : 0,
    statusSnapshot: product ? product.status : "",
    imageSnapshot: product ? product.image : ""
  };
}

function getPartnerNameByIdMap_(partnerData) {
  const result = {};
  partnerData.rows.forEach(row => {
    const partnerId = String(row[partnerData.col.partnerId] || "").trim();
    if (!partnerId) return;
    result[partnerId] = String(row[partnerData.col.partnerName] || "").trim();
  });
  return result;
}

function getPartnerCatalogItemCountMap_(catalogData) {
  const result = {};
  catalogData.rows.forEach(row => {
    if (!parsePartnerBoolean_(row[catalogData.col.visible])) return;
    const linkId = String(row[catalogData.col.linkId] || "").trim();
    if (!linkId) return;
    result[linkId] = (Number(result[linkId]) || 0) + 1;
  });
  return result;
}

function isPartnerLinkExpired_(link) {
  const time = link && link.expiresAt
    ? new Date(link.expiresAt).getTime()
    : NaN;
  return isNaN(time) || time <= new Date().getTime();
}

function mapPartnerRequestRow_(requestData, row) {
  return {
    requestId: String(row[requestData.col.requestId] || "").trim(),
    partnerId: String(row[requestData.col.partnerId] || "").trim(),
    partnerNameSnapshot: String(row[requestData.col.partnerNameSnapshot] || "").trim(),
    contactName: String(row[requestData.col.contactName] || "").trim(),
    contactPhone: String(row[requestData.col.contactPhone] || "").trim(),
    contactEmail: String(row[requestData.col.contactEmail] || "").trim(),
    message: String(row[requestData.col.message] || "").trim(),
    status: String(row[requestData.col.status] || "").trim(),
    itemCount: Number(row[requestData.col.itemCount]) || 0,
    estimatedTotal: Number(row[requestData.col.estimatedTotal]) || 0,
    submittedAt: row[requestData.col.submittedAt] || ""
  };
}

function mapPartnerRequestItemRow_(itemData, row) {
  const requestId = String(row[itemData.col.requestId] || "").trim();
  const sortOrder = Number(row[itemData.col.sortOrder]) || 0;
  return {
    requestItemId: requestId + "-" + (sortOrder || "item"),
    requestId,
    productId: String(row[itemData.col.productId] || "").trim(),
    nameSnapshot: String(row[itemData.col.nameSnapshot] || "").trim(),
    priceSnapshot: Number(row[itemData.col.priceSnapshot]) || 0,
    qty: Number(row[itemData.col.qty]) || 0,
    partnerNote: String(row[itemData.col.partnerNote] || "").trim(),
    statusSnapshot: String(row[itemData.col.statusSnapshot] || "").trim(),
    imageSnapshot: String(row[itemData.col.imageSnapshot] || "").trim(),
    sortOrder
  };
}

function clampPartnerRequestLimit_(rawLimit) {
  const limit = Number(rawLimit || 50);
  if (!Number.isFinite(limit) || limit <= 0) return 50;
  return Math.min(Math.floor(limit), 100);
}

function clampPartnerRequestOffset_(rawOffset) {
  const offset = Number(rawOffset || 0);
  if (!Number.isFinite(offset) || offset <= 0) return 0;
  return Math.floor(offset);
}

function getPartnerRequestDateTime_(value) {
  const time = value ? new Date(value).getTime() : 0;
  return isNaN(time) ? 0 : time;
}

function getPartnerCatalogContext_(ss, rawToken) {
  const linkData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerLinks
  );
  const partnerData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partners
  );
  const catalogData = getSheetDataBySchema_(
    ss,
    PARTNER_SHEET_NAMES.partnerCatalogItems
  );
  const productSheet = ss.getSheetByName("Products");

  if (!productSheet) {
    throw new Error("Products sheet not found");
  }

  const tokenHash = hashPartnerToken_(rawToken);
  const link = findPartnerLinkByTokenHash_(linkData, tokenHash);

  if (!link) {
    throw new Error("Partner link not found");
  }

  validatePartnerLink_(link);

  const partner = findPartnerById_(partnerData, link.partnerId);
  if (!partner) {
    throw new Error("Partner not found or disabled");
  }

  const productVariantCol = getProductVariantColumnMap_(
    ensureProductsVariantHeaders_(ss).headers
  );

  return {
    link,
    partner,
    selectedProductIds: getVisiblePartnerCatalogProductIds_(
      catalogData,
      link.linkId
    ),
    productMap: getProductRowMap_(productSheet),
    productVariantCol
  };
}

function validatePartnerLink_(link) {
  const linkStatus = normalizePartnerStatus_(link.status);
  if (linkStatus === "expired") {
    throw new Error("Partner link expired");
  }
  if (linkStatus === "disabled") {
    throw new Error("Partner link disabled");
  }
  if (linkStatus !== "active") {
    throw new Error("Partner link disabled");
  }
  if (!link.expiresAt || new Date(link.expiresAt) <= new Date()) {
    throw new Error("Partner link expired");
  }
}

function parsePartnerRequestItems_(rawItems) {
  let parsed;
  try {
    parsed = JSON.parse(String(rawItems || "[]"));
  } catch (err) {
    throw new Error("Invalid request items");
  }

  if (!Array.isArray(parsed) || !parsed.length) {
    throw new Error("Invalid request items");
  }

  if (parsed.length > PARTNER_REQUEST_MAX_ITEMS) {
    throw new Error(
      "Partner request cannot contain more than " +
      PARTNER_REQUEST_MAX_ITEMS +
      " items"
    );
  }

  return parsed.map((item, index) => {
    const productId = String(item && item.productId || "").trim().toUpperCase();
    const rawQty =
      item && item.hasOwnProperty("qty")
        ? item.qty
        : 1;
    const qty = Number(rawQty);

    if (!productId) {
      throw new Error("Product ID is required at row " + (index + 1));
    }

    if (!Number.isInteger(qty) || qty <= 0) {
      throw new Error("Invalid quantity for " + productId);
    }

    return {
      productId,
      qty,
      partnerNote: item ? item.partnerNote : ""
    };
  });
}

function sanitizePartnerText_(value, maxLength) {
  const text = String(value || "")
    .replace(/[\u0000-\u001F\u007F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  return text.slice(0, maxLength);
}

function rollbackPartnerRequestCreate_(
  requestSheet,
  itemSheet,
  requestId,
  headerCreated,
  itemRowsCreated
) {
  if (itemRowsCreated > 0) {
    deleteRowsByFirstColumnValue_(
      itemSheet,
      requestId,
      itemRowsCreated
    );
  }

  if (headerCreated) {
    deleteRowsByFirstColumnValue_(requestSheet, requestId, 1);
  }
}

function parseFuturePartnerDate_(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) {
    throw new Error("Expiration date is required");
  }

  const date = new Date(value);
  if (isNaN(date.getTime())) {
    throw new Error("Invalid expiration date");
  }

  if (date <= new Date()) {
    throw new Error("Expiration date must be in the future");
  }

  return date;
}

function parsePartnerProductIds_(rawValue) {
  let parsed;
  try {
    parsed = JSON.parse(String(rawValue || "[]"));
  } catch (err) {
    throw new Error("Invalid productIds JSON");
  }

  if (!Array.isArray(parsed)) {
    throw new Error("productIds must be an array");
  }

  const seen = {};
  const productIds = [];

  parsed.forEach(value => {
    const productId = String(value || "").trim().toUpperCase();
    if (!productId || seen[productId]) return;
    seen[productId] = true;
    productIds.push(productId);
  });

  return productIds;
}

function getSheetDataBySchema_(ss, sheetName) {
  const expectedHeaders = PARTNER_SHEET_SCHEMAS[sheetName];
  if (!expectedHeaders) {
    throw new Error("Unknown partner sheet: " + sheetName);
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + " sheet not found");
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = expectedHeaders.length;
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(header => String(header || "").trim());

  const matches = expectedHeaders.every((header, index) =>
    headers[index] === header
  );
  if (!matches) {
    throw new Error(sheetName + " schema mismatch");
  }

  const rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];
  const col = {};
  headers.forEach((header, index) => {
    col[header] = index;
  });

  return {
    sheet,
    headers,
    rows,
    col
  };
}

function findPartnerById_(partnerData, partnerId) {
  const targetId = String(partnerId || "").trim();
  const row = partnerData.rows.find(item =>
    String(item[partnerData.col.partnerId] || "").trim() === targetId &&
    normalizePartnerStatus_(item[partnerData.col.status]) === "active"
  );

  if (!row) return null;

  return {
    partnerId: String(row[partnerData.col.partnerId] || "").trim(),
    partnerName: String(row[partnerData.col.partnerName] || "").trim()
  };
}

function findPartnerLinkByTokenHash_(linkData, tokenHash) {
  const targetHash = String(tokenHash || "").trim();
  const row = linkData.rows.find(item =>
    String(item[linkData.col.tokenHash] || "").trim() === targetHash
  );

  if (!row) return null;

  return {
    linkId: String(row[linkData.col.linkId] || "").trim(),
    partnerId: String(row[linkData.col.partnerId] || "").trim(),
    status: String(row[linkData.col.status] || "").trim(),
    expiresAt: row[linkData.col.expiresAt],
    label: String(row[linkData.col.label] || "").trim()
  };
}

function getVisiblePartnerCatalogProductIds_(catalogData, linkId) {
  const targetLinkId = String(linkId || "").trim();
  return catalogData.rows
    .filter(row =>
      String(row[catalogData.col.linkId] || "").trim() === targetLinkId &&
      parsePartnerBoolean_(row[catalogData.col.visible])
    )
    .sort((left, right) =>
      (Number(left[catalogData.col.sortOrder]) || 0) -
      (Number(right[catalogData.col.sortOrder]) || 0)
    )
    .map(row => String(row[catalogData.col.productId] || "").trim().toUpperCase())
    .filter(Boolean);
}

function getProductRowMap_(productSheet) {
  const rows = productSheet.getDataRange().getValues();
  rows.shift();

  const productMap = new Map();
  rows.forEach(row => {
    const productId = String(row[0] || "").trim().toUpperCase();
    if (productId) {
      productMap.set(productId, row);
    }
  });

  return productMap;
}

function toSafePartnerProduct_(row, variantCol) {
  const stock = Number(row[3]) || 0;
  const status = stock <= 0
    ? "out"
    : String(row[6] || "").trim();

  return {
    productId: String(row[0] || "").trim(),
    name: String(row[1] || "").trim(),
    price: Number(row[2]) || 0,
    image:
      typeof row[4] === "string" && row[4].startsWith("http")
        ? row[4].trim()
        : "",
    status,
    detailsText: String(row[10] || "").trim(),
    compareImages: String(row[11] || "").trim(),
    ...getProductVariantMetadata_(row, variantCol)
  };
}

function isProductActive_(row) {
  const active = row[5];
  return (
    active === true ||
    active === "TRUE" ||
    active === 1 ||
    active === "1"
  );
}

function parsePartnerBoolean_(value) {
  return (
    value === true ||
    value === "TRUE" ||
    value === "true" ||
    value === 1 ||
    value === "1"
  );
}

function normalizePartnerStatus_(value) {
  return String(value || "").trim().toLowerCase();
}

function createPartnerRawToken_() {
  return (
    Utilities.getUuid().replace(/-/g, "") +
    Utilities.getUuid().replace(/-/g, "")
  );
}

function hashPartnerToken_(token) {
  return sha256Hex("partner:" + String(token || "").trim());
}

function getPartnerCatalogUrl_(rawToken) {
  let baseUrl = "";
  try {
    baseUrl = ScriptApp.getService().getUrl();
  } catch (err) {
    baseUrl = "";
  }

  const separator = baseUrl.indexOf("?") === -1 ? "?" : "&";
  return baseUrl
    ? baseUrl + separator + "partnerToken=" + encodeURIComponent(rawToken)
    : "?partnerToken=" + encodeURIComponent(rawToken);
}

function rollbackPartnerLinkCreate_(
  linkSheet,
  catalogSheet,
  linkId,
  linkCreated,
  catalogRowsCreated
) {
  if (catalogRowsCreated > 0) {
    deleteRowsByFirstColumnValue_(
      catalogSheet,
      linkId,
      catalogRowsCreated
    );
  }

  if (linkCreated) {
    deleteRowsByFirstColumnValue_(linkSheet, linkId, 1);
  }
}

function deleteRowsByFirstColumnValue_(sheet, value, maxRows) {
  const target = String(value || "").trim();
  if (!sheet || !target) return;

  const rows = sheet.getDataRange().getValues();
  let deleted = 0;

  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0] || "").trim() === target) {
      sheet.deleteRow(i + 1);
      deleted++;
      if (maxRows && deleted >= maxRows) return;
    }
  }
}

function getPendingDeliverySheet() {
  const sheet =
    getSS().getSheetByName("pending_delivery");

  if (!sheet) {
    throw new Error(
      "pending_delivery sheet not found"
    );
  }

  return sheet;
}

function createPendingDelivery(data, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  const sheet = getPendingDeliverySheet();

  /* ================= VALIDATE ================= */
  if (!data) {
    throw new Error("Missing pending delivery data");
  }

  const orderId =
    String(data.orderId || "").trim();

  const poNumber =
    String(data.poNumber || "").trim();

  const productId =
    String(data.productId || "").trim();

  const qty =
    Number(data.qty);

  const note =
    String(data.note || "").trim();

  if (
    !orderId ||
    !poNumber ||
    !productId ||
    !Number.isInteger(qty) ||
    qty <= 0
  ) {
    throw new Error(
      "Invalid pending delivery data"
    );
  }

  /* ================= VERIFY ORDER ================= */
  const orderSheet = getSS().getSheetByName("Orders");

  if (!orderSheet) {
    throw new Error("Orders sheet not found");
  }

  const orderRows = orderSheet.getDataRange().getValues();
  orderRows.shift();

  const orderRow = orderRows.find(
    row => String(row[0]).trim() === orderId
  );

  if (!orderRow) {
    throw new Error("Order not found");
  }

  const orderStatus =
    String(orderRow[3] || "").trim().toUpperCase();

  if (orderStatus !== "APPROVED") {
    throw new Error("Order is not approved");
  }

  const orderPoNumber =
    String(orderRow[7] || "").trim();

  if (orderPoNumber !== poNumber) {
    throw new Error("PO number does not match order");
  }

  const orderItems = JSON.parse(orderRow[1] || "[]");

  if (!Array.isArray(orderItems)) {
    throw new Error("Invalid order items");
  }

  const orderItem = orderItems.find(
    item => String(item.productId || "").trim() === productId
  );

  if (!orderItem) {
    throw new Error("Product not found in order");
  }

  const orderedQty = Number(orderItem.qty);

  if (
    !Number.isInteger(orderedQty) ||
    orderedQty <= 0
  ) {
    throw new Error("Invalid ordered quantity");
  }

  if (qty > orderedQty) {
    throw new Error("Pending quantity exceeds ordered quantity");
  }

 /* ================= DUPLICATE CHECK ================= */

  const rows =
    sheet.getDataRange().getValues();

  const headers =
    rows.shift();

  const orderIdx =
    headers.indexOf("orderId");

  const productIdx =
    headers.indexOf("productId");

  const statusIdx =
    headers.indexOf("status");

  const qtyIdx =
    headers.indexOf("qty");

  if (
    orderIdx === -1 ||
    productIdx === -1 ||
    statusIdx === -1 ||
    qtyIdx === -1
  ) {
    throw new Error(
      "pending_delivery schema mismatch"
    );
  }

  const cumulativeExistingQty = rows.reduce((sum, row) => {
    if (
      String(row[orderIdx]).trim() !== orderId ||
      String(row[productIdx]).trim() !== productId
    ) {
      return sum;
    }

    const existingQty = Number(row[qtyIdx]);
    if (!Number.isInteger(existingQty) || existingQty <= 0) {
      throw new Error("Invalid existing pending delivery quantity");
    }

    return sum + existingQty;
  }, 0);

  if (cumulativeExistingQty + qty > orderedQty) {
    throw new Error(
      "Cumulative pending quantity exceeds ordered quantity"
    );
  }

  const exists =
    rows.some(row =>
      String(row[orderIdx]).trim() === orderId &&
      String(row[productIdx]).trim() === productId &&
      String(row[statusIdx]).trim().toUpperCase() === "OPEN"
    );

  if (exists) {
    throw new Error(
      "รายการค้างส่งนี้มีอยู่แล้ว"
    );
  }

  /* ================= CREATE ================= */
  const pendingId =
    "PD-" + Date.now();

  const createdAt =
    new Date();

  sheet.appendRow([
    pendingId,     // pendingId
    orderId,       // orderId
    poNumber,      // poNumber
    productId,     // productId
    orderedQty,
    qty,           // qty
    note,          // note
    "OPEN",        // status
    createdAt,     // createdAt
    by || "",      // createdBy
    "",            // closedAt
    ""             // closedBy
  ]);

  return {
    success: true,
    pendingId
  };
  } finally {
    lock.releaseLock();
  }
}

function getPendingDeliveries() {

  const sheet = getPendingDeliverySheet();

  const rows = sheet.getDataRange().getValues();

  if (rows.length < 2) {
    return [];
  }

  const headers =
    rows.shift().map(h => String(h || "").trim());

  return rows
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      return obj;
    })
    .filter(item => item.pendingId);
}

function closePendingDelivery(pendingId, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  const sheet = getPendingDeliverySheet();

  const rows = sheet.getDataRange().getValues();

  if (rows.length < 2) {
    throw new Error("No pending deliveries found");
  }

  const headers =
    rows[0].map(h => String(h || "").trim());

  const pendingIdCol =
    headers.indexOf("pendingId");

  const statusCol =
    headers.indexOf("status");

  const closedAtCol =
    headers.indexOf("closedAt");

  const closedByCol =
    headers.indexOf("closedBy");

  if (
    pendingIdCol === -1 ||
    statusCol === -1 ||
    closedAtCol === -1
  ) {
    throw new Error(
      "pending_delivery schema mismatch"
    );
  }

  const rowIndex = rows.findIndex(
    (row, index) =>
      index > 0 &&
      String(row[pendingIdCol]).trim() ===
      String(pendingId).trim()
  );

  if (rowIndex === -1) {
    throw new Error("Pending delivery not found");
  }

  const currentStatus =
    String(rows[rowIndex][statusCol] || "").trim();

  if (currentStatus === "CLOSED") {
    throw new Error(
      "Pending delivery already closed"
    );
  }

  const updatedRow = rows[rowIndex].slice();
  updatedRow[statusCol] = "CLOSED";
  updatedRow[closedAtCol] = new Date();

  if (closedByCol !== -1) {
    updatedRow[closedByCol] = by || "";
  }

  sheet
    .getRange(rowIndex + 1, 1, 1, updatedRow.length)
    .setValues([updatedRow]);

  return {
    success: true,
    pendingId
  };
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.action) {
      return json({ success: false, message: "Missing action" });
    }

    const action = String(e.parameter.action).trim();

    // 🔓 PUBLIC ONLY
    if (action === "products") {
      return json({
        success: true,
        data: getProducts()
      });
    }

    // ❌ admin ห้ามผ่าน GET
    return json({
      success: false,
      message: "Forbidden"
    });

  } catch (err) {
    return json({
      success: false,
      error: err.message
    });
  }
}



function getProducts() {
  // 🔒 FIX: Web App ต้องอ้างอิง Spreadsheet แบบชัดเจน
  const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";
  const variantHeaderInfo = ensureProductsVariantHeaders_(getSS());
  if (variantHeaderInfo.appended.length) {
    CacheService.getScriptCache().remove(PUBLIC_PRODUCTS_CACHE_KEY);
  }
  const rows = getCachedPublicProductRows();

  // ❌ ไม่มีข้อมูลจริง (มีแต่ header)
  if (rows.length < 2) {
    return [];
  }

  const headers = rows.shift().map(h => String(h || "").trim()); // ลบ header
  const variantCol = getProductVariantColumnMap_(headers);

  return rows
    .map(r => {
      const productId = String(r[0] || "").trim();
      if (!productId) return null; // ❌ กันแถวว่าง

      const name   = String(r[1] || "").trim();
      const price  = Number(r[2]) || 0;
      const stock  = Number(r[3]) || 0;

      // 🖼 image (column E)
      const image =
        typeof r[4] === "string" && r[4].startsWith("http")
          ? r[4].trim()
          : "";

      // 🔓 active (column F)
      const active =
        r[5] === true ||
        r[5] === "TRUE" ||
        r[5] === 1 ||
        r[5] === "1";

      if (!active) return null;

      // 🟢 status (column G)
      const status = String(r[6] || "").trim();
      return {
        productId,
        name,
        price,
        stock,
        image,
        active,
        status,
        detailsText: String(r[10] || "").trim(),     // column K
        compareImages: String(r[11] || "").trim(),   // column L
        ...getProductVariantMetadata_(r, variantCol)
      };
    })
    .filter(Boolean); // ✅ ตัด null (แถวพัง / ว่าง)
}

function getAdminProducts() {
  const ss = getSS();
  ensureProductsVariantHeaders_(ss);

  const sheet = ss.getSheetByName("Products");
  if (!sheet) {
    return [];
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) {
    return [];
  }

  const headers = rows.shift().map(h => String(h || "").trim());

  // 🔴 CHANGED: อ่านตามชื่อ header ก่อน เพื่อกันคอลัมน์เลื่อน
  const col = {
    productId: headers.indexOf("productId"),
    name: headers.indexOf("name"),
    price: headers.indexOf("price"),
    stock: headers.indexOf("stock"),
    image: headers.indexOf("image"),
    active: headers.indexOf("active"),
    status: headers.indexOf("status"),
    note: headers.indexOf("note"),
    detailsText: headers.indexOf("detailsText"),
    compareImages: headers.indexOf("compareImages"),
    costPrice: headers.indexOf("costPrice"),
    parentProductId: headers.indexOf("parentProductId"),
    variantType: headers.indexOf("variantType"),
    variantValue: headers.indexOf("variantValue"),
    displayName: headers.indexOf("displayName"),
    variantSortOrder: headers.indexOf("variantSortOrder")
  };

  return rows
    .map(r => {
      // 🔴 CHANGED: fallback เป็น index เดิม หาก header ยังไม่ตรง
      const productId = String(
        col.productId > -1 ? r[col.productId] : r[0]
      ).trim();
      if (!productId) return null;

      const rawImage = col.image > -1 ? r[col.image] : r[4];
      const rawActive = col.active > -1 ? r[col.active] : r[5];
      const rawCostPrice = col.costPrice > -1 ? r[col.costPrice] : r[12];

      return {
        productId,
        name: String(col.name > -1 ? r[col.name] : r[1] || "").trim(),
        price: Number(col.price > -1 ? r[col.price] : r[2]) || 0,
        stock: Number(col.stock > -1 ? r[col.stock] : r[3]) || 0,
        image:
          typeof rawImage === "string" && rawImage.startsWith("http")
            ? rawImage.trim()
            : "",
        active:
          rawActive === true ||
          rawActive === "TRUE" ||
          rawActive === 1 ||
          rawActive === "1",
        status: String(col.status > -1 ? r[col.status] : r[6] || "").trim(),
        note: String(col.note > -1 ? r[col.note] : r[9] || "").trim(),
        detailsText: String(
          col.detailsText > -1 ? r[col.detailsText] : r[10] || ""
        ).trim(),
        compareImages: String(
          col.compareImages > -1 ? r[col.compareImages] : r[11] || ""
        ).trim(),
        costPrice:
          rawCostPrice !== "" &&
          rawCostPrice !== null &&
          rawCostPrice !== undefined
            ? Number(rawCostPrice)
            : null,
        ...getProductVariantMetadata_(r, col)
      };
    })
    .filter(Boolean);
}


function doPost(e) {
  try {
    /* ================= BASIC GUARD ================= */
    if (!e || !e.parameter) {
      throw new Error("Missing request data");
    }

    const params = e.parameter;
    const action = String(params.action || "").trim();

    Logger.log(
      "ACTION: " + action +
      " | TIMESTAMP: " + new Date().toISOString()
    );

    if (!action) {
      throw new Error("Missing action");
    }

    if (action === "getPartnerCatalog") {
      return json(
        getPartnerCatalog(params.partnerToken)
      );
    }

    if (action === "submitPartnerRequest") {
      return json(
        submitPartnerRequest(params)
      );
    }

    /* =================================================
       🔓 PUBLIC ACTION (NO AUTH REQUIRED)
    ================================================= */

    if (action === "adminLogin") {
      return json(
        adminLogin(
          params.username,
          params.password
        )
      );
    }

    if (action === "checkDuplicatePo") {
      return json(
        checkDuplicatePo(params)
      );
    }

    if (action === "createOrder") {
      enforceCreateOrderBodySize(e);

      const orderData = {
        items: JSON.parse(params.items || "[]"),
        poNumber: String(params.poNumber || "").trim()
      };

      const requestedQtyByProduct =
        validateCreateOrderRequestStructure(orderData);

      enforceInvalidCanonicalOrderRateLimit(requestedQtyByProduct);

      try {
        preflightCreateOrderCanonical(requestedQtyByProduct);
        clearInvalidCanonicalOrderAttempts(requestedQtyByProduct);
      } catch (err) {
        recordInvalidCanonicalOrderAttempt(requestedQtyByProduct);
        throw err;
      }

      enforceCreateOrderRateLimit(orderData);

      const result = createOrder(orderData);

      return json({
        success: true,
        data: result
      });
    }

    /* =================================================
       🔐 AUTH REQUIRED (ALL BELOW NEED TOKEN)
    ================================================= */

    const token = String(params.token || "").trim();
    if (!token) {
      throw new Error("Missing token");
    }

    const auth = requireAuth(token); // ❗ invalid → throw

    /* ========= ADD PRODUCT ========= */
    if (action === "addProduct") {
      const result = addProduct(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "bulkAddProducts") {
      return json(bulkAddProducts(e, auth));
    }

    /* ========= ADMIN READ ========= */
    if (action === "orders") {
      return json({
        success: true,
        data: getOrders()
      });
    }

    if (action === "adminProducts") {
      return json({
        success: true,
        data: getAdminProducts()
      });
    }

    if (action === "ensurePartnerCatalogSheets") {
      const result = ensurePartnerCatalogSheets();
      return json({
        success: true,
        data: {
          sheets: result.sheets
        }
      });
    }


    if (action === "createPartnerLink") {
      return json(
        createPartnerLink(params, auth)
      );
    }

    if (action === "listPartnerRequests") {
      return json(
        listPartnerRequests(params, auth)
      );
    }

    if (action === "getPartnerRequestDetail") {
      return json(
        getPartnerRequestDetail(params, auth)
      );
    }

    if (action === "updatePartnerRequestStatus") {
      return json(
        updatePartnerRequestStatus(params, auth)
      );
    }

    if (action === "convertPartnerRequestToOrder") {
      return json(
        convertPartnerRequestToOrder(params, auth)
      );
    }

    if (action === "listPartnerLinks") {
      return json(
        listPartnerLinks(params, auth)
      );
    }

    if (action === "getPartnerLinkDetail") {
      return json(
        getPartnerLinkDetail(params, auth)
      );
    }

    if (action === "updatePartnerLinkStatus") {
      return json(
        updatePartnerLinkStatus(params, auth)
      );
    }

    if (action === "stockLogs") {
      return json({
        success: true,
        data: getStockLogs()
      });
    }

    if (action === "pendingDeliveries") {
      return json({
        success: true,
        data: getPendingDeliveries()
      });
    }

    /* ========= ADMIN WRITE ========= */

    if (action === "approveOrder") {
      const result = approveOrder(
        token,
        params.orderId
      );

      return json({ success: true, data: result });
    }

    if (action === "rejectOrder") {
      const result = rejectOrder(token, params.orderId);
      return json({ success: true, data: result });
    }

    if (action === "stockIn") {
      const result = stockIn(token, {
        productId: params.productId,
        qty: Number(params.qty),
        reason: params.reason || ""
      });
      return json({ success: true, data: result });
    }

    if (action === "stockAdjust") {
      const result = stockAdjust(token, {
        productId: params.productId,
        newQty: Number(params.newQty),
        reason: params.reason || ""
      });
      return json({ success: true, data: result });
    }

    if (action === "createPendingDelivery") {
      const result = createPendingDelivery(
        {
          orderId: params.orderId,
          poNumber: params.poNumber,
          productId: params.productId,
          orderedQty: Number(params.orderedQty),
          qty: Number(params.qty),
          note: params.note || ""
        },
        auth.username
      );

      return json({
        success: true,
        data: result
      });
    }

    if (action === "closePendingDelivery") {
      const result = closePendingDelivery(
        params.pendingId,
        auth.username
      );

      return json({
        success: true,
        data: result
      });
    }

    if (action === "updateProduct") {
      const result = updateProduct(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "bulkUpdateProducts") {
      return json(bulkUpdateProducts(e, auth));
    }

    if (action === "deleteProduct") {
      const result = deleteProduct(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "uploadProductImage") {
      const result = uploadProductImage(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "adminLogout") {
      const result = adminLogout(params.token);
      return json({ success: true, data: result });
    }

    if (action === "changePassword") {
      const result = changePassword(
        params.token,
        params.currentPassword,
        params.newPassword
      );
      return json({ success: true, data: result });
    }

    if (action === "generateOrderPDF") {
       const result = generateOrderPDF(token, params.orderId);
       return json({ success: true, data: result });
     }


    /* ========= MAINTENANCE ========= */
    if (action === "cleanupImages") {
      const dryRun =
        String(params.dryRun || "true") === "true";

      const result = cleanupUnusedImages({
        by: auth.username,
        dryRun
      });

      return json({
        success: true,
        data: result
      });
    }


    /* =================================================
       ❌ UNKNOWN ACTION
    ================================================= */
    throw new Error("Invalid action: " + action);

  } catch (err) {
    Logger.log("doPost ERROR: " + err.message);

    return json({
      success: false,
      error: err.message || "Server error"
    });
  }
}





function uploadProductImage(e, auth) {
  /* ================= AUTH GUARD ================= */
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const by = auth.username;

  /* ================= PARAM GUARD ================= */
  if (!e || !e.parameter) {
    throw new Error("Invalid request (no parameters)");
  }

  /* ================= READ DATA ================= */
  let data = e.parameter.data;
  if (!data || typeof data !== "string") {
    throw new Error("No image data received");
  }

  const filename =
    String(e.parameter.filename || `product-${Date.now()}.jpg`).trim();

  const mimeType =
    String(e.parameter.mimeType || "image/jpeg").trim();

  /* ================= CLEAN BASE64 ================= */
  if (data.includes("base64,")) {
    data = data.split("base64,")[1];
  }

  /* ================= SIZE LIMIT GUARD ================= */
  const MAX_BASE64_SIZE = 5 * 1024 * 1024;
  if (data.length > MAX_BASE64_SIZE) {
    throw new Error("Image too large (max 5MB)");
  }

  /* ================= DECODE ================= */
  const bytes = Utilities.base64Decode(data);
  if (!bytes || bytes.length === 0) {
    throw new Error("Base64 decode failed");
  }

  const blob = Utilities.newBlob(bytes, mimeType, filename);

  /* ================= DRIVE ================= */
  const FOLDER_ID = "1un_A6DFFnknmEjx7LgACRKT7l8AgBBdK";
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // 🔴 CHANGED: เก็บ reference เพื่อ rollback ถ้า share ไม่ผ่าน
  let file = null;

  try {
    file = folder.createFile(blob);
    file.setName(filename);

    const fileId = file.getId();
    const imageUrl = `https://lh3.googleusercontent.com/d/${fileId}=w800`;

    Logger.log(
      `Image uploaded by ${by}: ${filename} (${fileId})`
    );

    return {
      imageUrl,
      uploadedBy: by
    };

  } catch (err) {
    // 🔴 CHANGED: ป้องกัน orphan file ถ้า create สำเร็จแต่ share พัง
    if (file) {
      try {
        file.setTrashed(true);
      } catch (_) {}
    }

    const message = String(err && err.message ? err.message : err);

    // 🔴 CHANGED: โยน error ให้ตรงสาเหตุจริงมากขึ้น
    if (
      message.includes("Access denied") ||
      message.includes("DriveApp")
    ) {
      throw new Error(
        "อัปโหลดรูปไม่สำเร็จ: ระบบไม่สามารถเข้าถึง Google Drive ได้"
      );
    }

    throw err;
  }
}


function stockAdjust(token, data) {
  // 🔐 AUTH (หัวใจ)
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const productSheet = ss.getSheetByName("Products");
  const logSheet = ss.getSheetByName("stock_logs");

  /* ================= VALIDATE INPUT ================= */
  if (
    !data ||
    !data.productId ||
    !Number.isInteger(data.newQty) ||
    data.newQty < 0
  ) {
    throw new Error("Invalid stock adjust data");
  }

  /* ================= LOAD PRODUCTS ================= */
  const products = productSheet.getDataRange().getValues();
  products.shift(); // remove header

  const idx = products.findIndex(
    r => String(r[0]).trim() === data.productId
  );

  if (idx === -1) {
    throw new Error("Product not found");
  }

  const before = Number(products[idx][3]);
  const after  = Number(data.newQty);
  const diff   = after - before;

  /* ================= UPDATE STOCK ================= */
  productSheet
    .getRange(idx + 2, 4) // column D: stock
    .setValue(after);

  /* ================= LOG ADJUST ================= */
  try {
    logSheet.appendRow([
      "LOG-" + Date.now(), // logId
      data.productId,      // productId
      "ADJUST",            // type
      diff,                // qty (delta)
      before,              // before
      after,               // after
      by,                  // by (จาก token)
      "",                  // orderId
      data.reason || "",   // reason
      new Date()           // timestamp
    ]);
  } catch (err) {
    try {
      productSheet
        .getRange(idx + 2, 4)
        .setValue(before);
    } catch (rollbackErr) {
      throw new Error(
        String(err.message || err) +
        " | Stock rollback failed: " +
        String(rollbackErr.message || rollbackErr)
      );
    }

    throw err;
  }

  return {
    success: true,
    adjustedBy: by
  };
  } finally {
    lock.releaseLock();
  }
}




const CREATE_ORDER_MAX_ITEMS = 50;
const CREATE_ORDER_MAX_QTY_PER_PRODUCT = 999;
const CREATE_ORDER_MAX_BODY_BYTES = 100 * 1024;
const CREATE_ORDER_MAX_PO_LENGTH = 80;
const CREATE_ORDER_RATE_LIMIT_MAX_REQUESTS_PER_BUCKET = 5;
const CREATE_ORDER_RATE_LIMIT_GLOBAL_MAX_REQUESTS = 100;
const CREATE_ORDER_RATE_LIMIT_WINDOW_MS = 2 * 60 * 1000;
const CREATE_ORDER_RATE_LIMIT_KEY = "CREATE_ORDER_RATE_LIMIT";
const PUBLIC_PRODUCTS_CACHE_KEY = "PUBLIC_PRODUCTS_ROWS";
const PUBLIC_PRODUCTS_CACHE_SECONDS = 10;
const INVALID_CANONICAL_ORDER_MAX_ATTEMPTS = 5;
const INVALID_CANONICAL_ORDER_WINDOW_MS = 2 * 60 * 1000;
const INVALID_CANONICAL_ORDER_RATE_LIMIT_PREFIX =
  "INVALID_CANONICAL_ORDER_";

function getCachedPublicProductRows() {
  const cache = CacheService.getScriptCache();
  const cachedRows = cache.get(PUBLIC_PRODUCTS_CACHE_KEY);

  if (cachedRows) {
    try {
      const rows = JSON.parse(cachedRows);
      if (Array.isArray(rows)) {
        return rows;
      }
    } catch (err) {
      cache.remove(PUBLIC_PRODUCTS_CACHE_KEY);
    }
  }

  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Products");

  if (!sheet) {
    return [];
  }

  const rows = sheet.getDataRange().getValues();
  try {
    cache.put(
      PUBLIC_PRODUCTS_CACHE_KEY,
      JSON.stringify(rows),
      PUBLIC_PRODUCTS_CACHE_SECONDS
    );
  } catch (err) {
    Logger.log("Public products cache skipped: " + err.message);
  }
  return rows;
}

function getInvalidCanonicalOrderRateLimitKey(requestedQtyByProduct) {
  const signature = Array.from(requestedQtyByProduct.keys())
    .sort()
    .map(productId =>
      productId + ":" + requestedQtyByProduct.get(productId)
    )
    .join("|");
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    signature,
    Utilities.Charset.UTF_8
  );
  const hash = digest
    .slice(0, 12)
    .map(byte => (byte + 256).toString(16).slice(-2))
    .join("");

  return INVALID_CANONICAL_ORDER_RATE_LIMIT_PREFIX + hash;
}

function getInvalidCanonicalOrderAttemptState(requestedQtyByProduct) {
  const key =
    getInvalidCanonicalOrderRateLimitKey(requestedQtyByProduct);
  const cache = CacheService.getScriptCache();
  const rawState = cache.get(key);

  if (!rawState) {
    return { key, cache, state: null };
  }

  try {
    return {
      key,
      cache,
      state: JSON.parse(rawState)
    };
  } catch (err) {
    cache.remove(key);
    return { key, cache, state: null };
  }
}

function enforceInvalidCanonicalOrderRateLimit(
  requestedQtyByProduct
) {
  const attempt =
    getInvalidCanonicalOrderAttemptState(requestedQtyByProduct);
  const state = attempt.state;

  if (!state) {
    return;
  }

  const startedAt = Number(state.startedAt);
  const count = Number(state.count);

  if (
    !Number.isFinite(startedAt) ||
    !Number.isInteger(count) ||
    Date.now() - startedAt >= INVALID_CANONICAL_ORDER_WINDOW_MS
  ) {
    attempt.cache.remove(attempt.key);
    return;
  }

  if (count >= INVALID_CANONICAL_ORDER_MAX_ATTEMPTS) {
    throw new Error(
      "Too many invalid order requests. Please try again later"
    );
  }
}

function recordInvalidCanonicalOrderAttempt(
  requestedQtyByProduct
) {
  const attempt =
    getInvalidCanonicalOrderAttemptState(requestedQtyByProduct);
  const now = Date.now();
  let state = attempt.state;

  if (
    !state ||
    !Number.isFinite(Number(state.startedAt)) ||
    !Number.isInteger(Number(state.count)) ||
    now - Number(state.startedAt) >=
      INVALID_CANONICAL_ORDER_WINDOW_MS
  ) {
    state = {
      count: 0,
      startedAt: now
    };
  }

  state.count = Number(state.count) + 1;
  attempt.cache.put(
    attempt.key,
    JSON.stringify(state),
    Math.ceil(INVALID_CANONICAL_ORDER_WINDOW_MS / 1000)
  );
}

function clearInvalidCanonicalOrderAttempts(
  requestedQtyByProduct
) {
  CacheService
    .getScriptCache()
    .remove(
      getInvalidCanonicalOrderRateLimitKey(
        requestedQtyByProduct
      )
    );
}

function enforceCreateOrderBodySize(e) {
  if (!e || !e.postData) {
    return;
  }

  let bodySize = Number(e.postData.length);

  if (!Number.isFinite(bodySize)) {
    bodySize = Utilities
      .newBlob(String(e.postData.contents || ""))
      .getBytes()
      .length;
  }

  if (bodySize > CREATE_ORDER_MAX_BODY_BYTES) {
    throw new Error("Order request is too large");
  }
}

function validateCreateOrderRequestStructure(data) {
  if (!data || !Array.isArray(data.items)) {
    throw new Error("Invalid order items");
  }

  if (data.items.length === 0) {
    throw new Error("Order must contain at least 1 item");
  }

  if (data.items.length > CREATE_ORDER_MAX_ITEMS) {
    throw new Error(
      `Order cannot contain more than ${CREATE_ORDER_MAX_ITEMS} items`
    );
  }

  const poNumber = String(data.poNumber || "").trim();

  if (!poNumber) {
    throw new Error("PO number is required");
  }

  if (poNumber.length > CREATE_ORDER_MAX_PO_LENGTH) {
    throw new Error(
      `PO number cannot exceed ${CREATE_ORDER_MAX_PO_LENGTH} characters`
    );
  }

  if (/^[=+\-@]/.test(poNumber)) {
    throw new Error("PO number contains an unsafe leading character");
  }

  const requestedQtyByProduct = new Map();

  data.items.forEach((item, index) => {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      throw new Error(`Invalid item at index ${index}`);
    }

    const productId = String(item.productId || "").trim();
    const qty = Number(item.qty);

    if (
      !productId ||
      !Number.isInteger(qty) ||
      qty <= 0 ||
      qty > CREATE_ORDER_MAX_QTY_PER_PRODUCT
    ) {
      throw new Error(`Invalid item at index ${index}`);
    }

    const mergedQty =
      (requestedQtyByProduct.get(productId) || 0) + qty;

    if (mergedQty > CREATE_ORDER_MAX_QTY_PER_PRODUCT) {
      throw new Error(`Quantity exceeds limit for ${productId}`);
    }

    requestedQtyByProduct.set(productId, mergedQty);
  });

  return requestedQtyByProduct;
}

function preflightCreateOrderCanonical(requestedQtyByProduct) {
  const productRows = getCachedPublicProductRows();
  productRows.shift();

  requestedQtyByProduct.forEach((qty, productId) => {
    const productRow = productRows.find(
      row => String(row[0] || "").trim() === productId
    );

    if (!productRow) {
      throw new Error("Product not found: " + productId);
    }

    const currentStock = Number(productRow[3]);
    const rawActive = productRow[5];
    const active =
      rawActive === true ||
      rawActive === "TRUE" ||
      rawActive === 1 ||
      rawActive === "1";

    if (!active) {
      throw new Error("Product is not active: " + productId);
    }

    if (!Number.isFinite(currentStock) || currentStock < qty) {
      throw new Error(
        `Stock not enough for ${productId} (remain ${currentStock})`
      );
    }
  });
}

function getCreateOrderRateLimitBucketKey(data) {
  const poNumber = String(data && data.poNumber || "")
    .trim()
    .toLowerCase();
  const itemTotals = new Map();
  const items = data && Array.isArray(data.items) ? data.items : [];

  items.forEach(item => {
    const productId = String(item && item.productId || "")
      .trim()
      .toLowerCase();
    const qty = Number(item && item.qty);
    const itemKey = productId || "[missing-product]";
    const normalizedQty = Number.isFinite(qty) ? qty : "[invalid-qty]";
    itemTotals.set(
      itemKey,
      Number(itemTotals.get(itemKey) || 0) + normalizedQty
    );
  });

  const itemSignature = Array.from(itemTotals.keys())
    .sort()
    .map(productId => productId + ":" + itemTotals.get(productId))
    .join("|");
  const signature = poNumber + "|" + itemSignature;
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    signature,
    Utilities.Charset.UTF_8
  );

  return digest
    .slice(0, 12)
    .map(byte => (byte + 256).toString(16).slice(-2))
    .join("");
}

function enforceCreateOrderRateLimit(data) {
    const properties = PropertiesService.getScriptProperties();
    const rawState = properties.getProperty(CREATE_ORDER_RATE_LIMIT_KEY);
    const now = Date.now();
    const bucketKey = getCreateOrderRateLimitBucketKey(data);
    let state = {};

    if (rawState) {
      try {
        state = JSON.parse(rawState);
      } catch (err) {
        state = {};
      }
    }

    if (!state || typeof state !== "object") {
      state = {};
    }

    if (!state.buckets || typeof state.buckets !== "object") {
      state.buckets = {};
    }

    Object.keys(state.buckets).forEach(key => {
      const bucket = state.buckets[key];
      const bucketStartedAt =
        Array.isArray(bucket) ? Number(bucket[1]) : NaN;

      if (
        !Number.isFinite(bucketStartedAt) ||
        now - bucketStartedAt >= CREATE_ORDER_RATE_LIMIT_WINDOW_MS
      ) {
        delete state.buckets[key];
      }
    });

    let globalState = Array.isArray(state.global)
      ? state.global
      : [0, now];

    if (
      !Number.isInteger(Number(globalState[0])) ||
      Number(globalState[0]) < 0 ||
      !Number.isFinite(Number(globalState[1])) ||
      now - Number(globalState[1]) >= CREATE_ORDER_RATE_LIMIT_WINDOW_MS
    ) {
      globalState = [0, now];
    }

    let bucketState = Array.isArray(state.buckets[bucketKey])
      ? state.buckets[bucketKey]
      : [0, now];

    if (
      !Number.isInteger(Number(bucketState[0])) ||
      Number(bucketState[0]) < 0 ||
      !Number.isFinite(Number(bucketState[1]))
    ) {
      bucketState = [0, now];
    }

    if (
      Number(bucketState[0]) >=
      CREATE_ORDER_RATE_LIMIT_MAX_REQUESTS_PER_BUCKET
    ) {
      throw new Error("Too many order requests. Please try again later");
    }

    if (
      Number(globalState[0]) >=
      CREATE_ORDER_RATE_LIMIT_GLOBAL_MAX_REQUESTS
    ) {
      throw new Error("Too many order requests. Please try again later");
    }

    bucketState[0] = Number(bucketState[0]) + 1;
    globalState[0] = Number(globalState[0]) + 1;
    state.buckets[bucketKey] = bucketState;
    state.global = globalState;

    properties.setProperty(
      CREATE_ORDER_RATE_LIMIT_KEY,
      JSON.stringify(state)
    );
}

function createOrder(data) {
  // 🔒 FIX: Web App ต้องอ้างอิง Spreadsheet แบบชัดเจน
  /* ================= VALIDATE ROOT ================= */
  if (!data || !Array.isArray(data.items)) {
    throw new Error("Invalid order items");
  }

  if (data.items.length === 0) {
    throw new Error("Order must contain at least 1 item");
  }

  if (data.items.length > CREATE_ORDER_MAX_ITEMS) {
    throw new Error(
      `Order cannot contain more than ${CREATE_ORDER_MAX_ITEMS} items`
    );
  }

  const poNumber = String(data.poNumber || "").trim();

  if (!poNumber) {
    throw new Error("PO number is required");
  }

  if (poNumber.length > CREATE_ORDER_MAX_PO_LENGTH) {
    throw new Error(
      `PO number cannot exceed ${CREATE_ORDER_MAX_PO_LENGTH} characters`
    );
  }

  if (/^[=+\-@]/.test(poNumber)) {
    throw new Error("PO number contains an unsafe leading character");
  }

  /* ================= MERGE CLIENT ITEMS ================= */
  const requestedQtyByProduct = new Map();

  data.items.forEach((item, index) => {
    const productId = String(item.productId || "").trim();
    const qty       = Number(item.qty);

    if (
      !productId ||
      !Number.isInteger(qty) ||
      qty <= 0 ||
      qty > CREATE_ORDER_MAX_QTY_PER_PRODUCT
    ) {
      throw new Error(
        `Invalid item at index ${index}`
      );
    }

    const mergedQty =
      (requestedQtyByProduct.get(productId) || 0) + qty;

    if (mergedQty > CREATE_ORDER_MAX_QTY_PER_PRODUCT) {
      throw new Error(
        `Quantity exceeds limit for ${productId}`
      );
    }

    requestedQtyByProduct.set(productId, mergedQty);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Orders");

  if (!sheet) {
    throw new Error("Orders sheet not found");
  }

  ensureOrdersExtendedHeaders_(ss);

  /* ================= LOAD AUTHORITATIVE PRODUCTS ================= */
  const productSheet = ss.getSheetByName("Products");
  if (!productSheet) {
    throw new Error("Products sheet not found");
  }

  const productRows = productSheet.getDataRange().getValues();
  productRows.shift(); // remove header

  const cleanItems = [];
  let total = 0;

  requestedQtyByProduct.forEach((qty, productId) => {
    const pIndex = productRows.findIndex(
      r => String(r[0]).trim() === productId
    );

    if (pIndex === -1) {
      throw new Error("Product not found: " + productId);
    }

    const productRow = productRows[pIndex];
    const name = String(productRow[1] || "").trim();
    const price = Number(productRow[2]);
    const costPrice = Number(productRow[12]);
    const currentStock = Number(productRow[3]);
    const rawActive = productRow[5];
    const active =
      rawActive === true ||
      rawActive === "TRUE" ||
      rawActive === 1 ||
      rawActive === "1";

    if (
      !name ||
      !Number.isFinite(price) ||
      price < 0 ||
      !Number.isFinite(costPrice) ||
      costPrice < 0
    ) {
      throw new Error("Invalid product data: " + productId);
    }

    if (!active) {
      throw new Error("Product is not active: " + productId);
    }

    if (!Number.isFinite(currentStock) || currentStock < qty) {
      throw new Error(
        `Stock not enough for ${productId} (remain ${currentStock})`
      );
    }

    cleanItems.push({
      productId,
      name,
      qty,
      price,
      costPrice
    });

    total += qty * price;
  });

  /* ================= CREATE ORDER ================= */
  const orderId   = "ORD-" + Utilities.getUuid();
  const status    = "PENDING";
  const createdAt = new Date();

  sheet.appendRow([
    orderId,
    JSON.stringify(cleanItems),
    total,
    status,
    createdAt,
    "",          // approvedAt
    "",          // approvedBy
    poNumber     // poNumber
  ]);

  return {
    success: true,
    orderId,
    total,
    items: cleanItems
  };
  } finally {
    lock.releaseLock();
  }
}



function json(payload) {
  // 🔒 normalize null / undefined
  if (payload === null || payload === undefined) {
    payload = {};
  }

  // 🔴 ห้ามส่ง array ตรง ๆ (กันพังเงียบ)
  if (Array.isArray(payload)) {
    throw new Error("json() payload must be an object, not array");
  }

  // 🔒 enforce success flag (ค่าเดียว เชื่อถือได้)
  const success =
    payload.success === false || payload.error ? false : true;

  // ❌ ห้ามให้ success ซ้อน
  const { success: _ignored, ...rest } = payload;

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success,
        ...rest
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}




function approveOrder(token, orderId) {
  /* ================= AUTH (GLOBAL) ================= */
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    /* ================= SPREADSHEET ================= */
    const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const orderSheet   = ss.getSheetByName("Orders");
    const productSheet = ss.getSheetByName("Products");
    const logSheet     = ss.getSheetByName("stock_logs");

    /* ================= LOAD ORDER ================= */
    const orders = orderSheet.getDataRange().getValues();
    orders.shift();

    const idx = orders.findIndex(r => r[0] === orderId);
    if (idx === -1) {
      throw new Error("Order not found");
    }

    const rowNumber = idx + 2;
    const orderRow = orders[idx];
    const status = String(orderRow[3] || "").trim().toUpperCase();

    if (status === "APPROVED") {
      return {
        success: true,
        approvedBy: String(orderRow[6] || by)
      };
    }

    if (status === "REJECTED") {
      throw new Error("Order already rejected");
    }

    if (status !== "PENDING") {
      throw new Error("Invalid order status");
    }

    const items = JSON.parse(orderRow[1] || "[]");
    if (!Array.isArray(items) || items.length === 0) {
      throw new Error("Invalid order items");
    }

    /* ================= LOAD PRODUCTS ================= */
    const products = productSheet.getDataRange().getValues();
    products.shift();

    /* ================= MERGE + VALIDATE STOCK ================= */
    const requiredQtyByProduct = new Map();

    items.forEach((item, itemIndex) => {
      const productId = String(item.productId || "").trim();
      const qty = Number(item.qty);

      if (
        !productId ||
        !Number.isInteger(qty) ||
        qty <= 0
      ) {
        throw new Error("Invalid order item at index " + itemIndex);
      }

      requiredQtyByProduct.set(
        productId,
        (requiredQtyByProduct.get(productId) || 0) + qty
      );
    });

    const mutations = [];

    requiredQtyByProduct.forEach((qty, productId) => {
      const pIndex = products.findIndex(
        r => String(r[0]).trim() === productId
      );

      if (pIndex === -1) {
        throw new Error("Product not found: " + productId);
      }

      const currentStock = Number(products[pIndex][3]);
      if (!Number.isFinite(currentStock) || currentStock < qty) {
        throw new Error(
          `Stock not enough for ${productId} (remain ${currentStock})`
        );
      }

      mutations.push({
        productId,
        rowNumber: pIndex + 2,
        qty,
        before: currentStock,
        after: currentStock - qty
      });
    });

    /* ================= APPLY + COMPENSATE ON FAILURE ================= */
    const appliedMutations = [];
    const createdLogIds = [];

    try {
      mutations.forEach(mutation => {
        productSheet
          .getRange(mutation.rowNumber, 4)
          .setValue(mutation.after);
        appliedMutations.push(mutation);
      });

      const logTimestamp = new Date();

      mutations.forEach(mutation => {
        const logId = "LOG-" + Utilities.getUuid();
        createdLogIds.push(logId);

        logSheet.appendRow([
          logId,
          mutation.productId,
          "OUT",
          mutation.qty,
          mutation.before,
          mutation.after,
          by,
          orderId,
          "",
          logTimestamp
        ]);
      });

      orderSheet
        .getRange(rowNumber, 4, 1, 4)
        .setValues([[
          "APPROVED",
          orderRow[4],
          new Date(),
          by
        ]]);
    } catch (err) {
      let rollbackError = null;

      if (createdLogIds.length > 0) {
        try {
          const createdLogIdSet = new Set(createdLogIds);
          const lastLogRow = logSheet.getLastRow();

          if (lastLogRow >= 2) {
            const logIds = logSheet
              .getRange(2, 1, lastLogRow - 1, 1)
              .getValues();

            const rowsToDelete = [];

            logIds.forEach((row, index) => {
              if (createdLogIdSet.has(String(row[0]))) {
                rowsToDelete.push(index + 2);
              }
            });

            rowsToDelete
              .sort((a, b) => b - a)
              .forEach(logRowNumber => {
                logSheet.deleteRow(logRowNumber);
              });
          }
        } catch (logRollbackErr) {
          rollbackError = logRollbackErr;
        }
      }

      if (appliedMutations.length > 0) {
        try {
          appliedMutations.reverse().forEach(mutation => {
            productSheet
              .getRange(mutation.rowNumber, 4)
              .setValue(mutation.before);
          });
        } catch (stockRollbackErr) {
          rollbackError = rollbackError || stockRollbackErr;
        }
      }

      if (rollbackError) {
        throw new Error(
          String(err.message || err) +
          " | Rollback failed: " +
          String(rollbackError.message || rollbackError)
        );
      }

      throw err;
    }

    return {
      success: true,
      approvedBy: by
    };
  } finally {
    lock.releaseLock();
  }
}




function rejectOrder(token, orderId) {
  // 🔐 AUTH
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const orderSheet = ss.getSheetByName("Orders");

  const allRows = orderSheet.getDataRange().getValues();
  const rows = allRows.slice(1);

  const orderRowIndex = rows.findIndex(r => r[0] === orderId);
  if (orderRowIndex === -1) {
    throw new Error("Order not found");
  }

  const rowNumber = orderRowIndex + 2;
  const status = rows[orderRowIndex][3];

  // ❌ ป้องกัน reject ซ้ำ
  if (status !== "PENDING") {
    throw new Error("Order already processed");
  }

  // ===== Update status + audit =====
  const updatedRow = rows[orderRowIndex].slice();
  updatedRow[3] = "REJECTED";
  updatedRow[5] = new Date();
  updatedRow[6] = by;

  orderSheet
    .getRange(rowNumber, 1, 1, updatedRow.length)
    .setValues([updatedRow]);

  return {
    success: true,
    rejectedBy: by
  };
  } finally {
    lock.releaseLock();
  }
}



 function getStockLogs() {
   // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
   const sheet = getSS().getSheetByName("stock_logs");

   if (!sheet) {
     return [];
   }

   const rows = sheet.getDataRange().getValues();
   if (rows.length < 2) return [];
  const headers = rows.shift().map(h => String(h || "").trim());

   return rows.map(r => {
     let obj = {};
     headers.forEach((h, i) => obj[h] = r[i]);

    // ✅ backward compatibility:
    // ชีตเก่าไม่มีคอลัมน์ reason และ IN/ADJUST เคยเอา reason ไปใส่ไว้ในช่อง orderId
    if (!("reason" in obj) || obj.reason === "" || obj.reason == null) {
      if (
        (obj.type === "IN" || obj.type === "ADJUST") &&
        obj.orderId &&
        !String(obj.orderId).startsWith("ORD-")
      ) {
        obj.reason = obj.orderId;
        obj.orderId = "";
      }
    }

     return obj;
   });
 }


function isProductReferencedByPendingOrder(ss, productId) {
  const orderSheet = ss.getSheetByName("Orders");
  if (!orderSheet) {
    throw new Error("Orders sheet not found");
  }

  const rows = orderSheet.getDataRange().getValues();
  if (rows.length < 2) {
    return false;
  }

  const headers = rows[0].map(header => String(header || "").trim());
  const itemsCol = headers.indexOf("items");
  const statusCol = headers.indexOf("status");

  if (itemsCol === -1 || statusCol === -1) {
    throw new Error("Orders schema mismatch");
  }

  const normalizedProductId = String(productId || "").trim();

  return rows.slice(1).some(row => {
    const status = String(row[statusCol] || "").trim().toUpperCase();
    if (status !== "PENDING") {
      return false;
    }

    let items;
    try {
      items = JSON.parse(row[itemsCol] || "[]");
    } catch (err) {
      throw new Error("Invalid pending order items");
    }

    if (!Array.isArray(items)) {
      throw new Error("Invalid pending order items");
    }

    return items.some(
      item =>
        String(item && item.productId || "").trim() ===
        normalizedProductId
    );
  });
}


function updateProduct(e, auth) {
  try {
    // 🔐 AUTH GUARD
    if (!auth || !auth.username) {
      throw new Error("Unauthorized");
    }

    const by = auth.username;

    /* ================= READ PARAM ================= */
    const oldProductId = String(e.parameter.oldProductId || "").trim();
    const newProductId = String(e.parameter.newProductId || "").trim();
    const name   = String(e.parameter.name || "").trim();
    const price  = Number(e.parameter.price);
    const rawCostPrice = String(e.parameter.costPrice ?? "").trim();
    const costPrice = rawCostPrice === "" ? null : Number(rawCostPrice);
    const stock  = Number(e.parameter.stock);
    const image  = String(e.parameter.image || "").trim();
    const status = String(e.parameter.status || "").trim();
    const note   = String(e.parameter.note || "").trim();
    const detailsText   = String(e.parameter.detailsText || "").trim();
    const compareImages = String(e.parameter.compareImages || "").trim();
    const variantMetadata = parseProductVariantMetadata_(e.parameter);


    /* ================= VALIDATE ================= */
    if (!oldProductId || !newProductId || !name ||
       !Number.isInteger(price) ||
       costPrice === null ||
       !Number.isInteger(costPrice) ||
       !Number.isInteger(stock)) {
      throw new Error("Invalid product data");
    }

    if (price < 0 || costPrice < 0 || stock < 0) {
      throw new Error("Price, cost price and stock must be >= 0");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const ss = getSS();
      const sh = ss.getSheetByName("Products");
      if (!sh) {
        throw new Error("Sheet Products not found");
      }
      const productHeaders = ensureProductsVariantHeaders_(ss).headers;

      const data = sh.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === oldProductId) {

          // 🔒 อ่าน active และ stock ล่าสุดภายใน lock
          const originalRow = data[i].slice();
          const currentActive = data[i][5]; // column F
          const oldStock = Number(data[i][3]);
          const stockChanged = oldStock !== stock;
          const logSheet = stockChanged
            ? ss.getSheetByName("stock_logs")
            : null;

          if (stockChanged && !logSheet) {
            throw new Error("Sheet stock_logs not found");
          }

          try {
            // ===== HANDLE SKU CHANGE =====
            if (oldProductId !== newProductId) {
              if (isProductReferencedByPendingOrder(ss, oldProductId)) {
                throw new Error(
                  "Cannot change SKU while it is referenced by a pending order"
                );
              }

              // 🔍 Duplicate guard
              const exists = data.slice(1).some(
                r => String(r[0]).trim() === newProductId
              );
              if (exists) {
                throw new Error("SKU already exists");
              }

              // 🔄 Update SKU column A
              sh.getRange(i + 1, 1).setValue(newProductId);
            }

            /* ================= UPDATE ================= */
            sh.getRange(i + 1, 2).setValue(name);          // B: name
            sh.getRange(i + 1, 3).setValue(price);         // C: price
            sh.getRange(i + 1, 4).setValue(stock);         // D: stock
            sh.getRange(i + 1, 5).setValue(image);         // E: image
            sh.getRange(i + 1, 6).setValue(currentActive); // F: active (คงเดิม)
            sh.getRange(i + 1, 7).setValue(status);        // G: status
            sh.getRange(i + 1, 10).setValue(note);         // J: note
            sh.getRange(i + 1, 11).setValue(detailsText);   // K: detailsText
            sh.getRange(i + 1, 12).setValue(compareImages); // L: compareImages
            sh.getRange(i + 1, 13).setValue(costPrice);     // M: costPrice
            writeProductVariantMetadata_(
              sh,
              i + 1,
              variantMetadata,
              productHeaders
            );

            if (stockChanged) {
              logSheet.appendRow([
                "LOG-" + Date.now(),
                newProductId,
                "ADJUST",
                stock - oldStock,
                oldStock,
                stock,
                by,
                "",
                "",
                new Date()
              ]);
            }
          } catch (err) {
            try {
              sh
                .getRange(i + 1, 1, 1, originalRow.length)
                .setValues([originalRow]);
            } catch (rollbackErr) {
              throw new Error(
                String(err.message || err) +
                " | Product rollback failed: " +
                String(rollbackErr.message || rollbackErr)
              );
            }

            throw err;
          }

          Logger.log(`Product ${newProductId} updated by ${by}`);

          return {
            updatedBy: by
          };
        }
      }

      throw new Error("ไม่พบสินค้า");
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log("updateProduct error: " + err);
    throw err;
  }
}

function deleteProduct(e, auth) {
  // 🔐 AUTH GUARD
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const by = auth.username;

  const productId = String(e.parameter.productId || "").trim();
  if (!productId) {
    throw new Error("Missing productId");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    const sh = ss.getSheetByName("Products");
    if (!sh) {
      throw new Error("Sheet Products not found");
    }

    const rows = sh.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === productId) {
        if (isProductReferencedByPendingOrder(ss, productId)) {
          throw new Error(
            "Cannot delete product while it is referenced by a pending order"
          );
        }

        // ❌ HARD DELETE: ลบแถวออกจากชีตจริง
        sh.deleteRow(i + 1);

        Logger.log(`Product ${productId} hard-deleted by ${by}`);

        return {
          success: true,
          deletedBy: by
        };
      }
    }

    throw new Error("Product not found");
  } finally {
    lock.releaseLock();
  }
}

function stockIn(token, data) {
  // 🔐 AUTH
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const productSheet = ss.getSheetByName("Products");
  const logSheet = ss.getSheetByName("stock_logs");

  // 🔍 Validate input
  if (!data.productId || !Number.isInteger(data.qty) || data.qty <= 0) {
    throw new Error("Invalid stock in data");
  }

  const products = productSheet.getDataRange().getValues();
  products.shift(); // header

  const idx = products.findIndex(r => r[0] === data.productId);
  if (idx === -1) {
    throw new Error("Product not found");
  }

  const before = Number(products[idx][3]);
  const after = before + data.qty;

  // ===== Update stock =====
  productSheet
    .getRange(idx + 2, 4) // column stock
    .setValue(after);

  // ===== Log IN =====
  try {
    logSheet.appendRow([
      "LOG-" + Date.now(),   // logId
      data.productId,        // productId
      "IN",                  // type
      data.qty,              // qty
      before,                // before
      after,                 // after
      by,                    // by (จาก token)
      "",                    // orderId
      data.reason || "",     // reason
      new Date()             // timestamp
    ]);
  } catch (err) {
    try {
      productSheet
        .getRange(idx + 2, 4)
        .setValue(before);
    } catch (rollbackErr) {
      throw new Error(
        String(err.message || err) +
        " | Stock rollback failed: " +
        String(rollbackErr.message || rollbackErr)
      );
    }

    throw err;
  }

  return {
    success: true,
    by
  };
  } finally {
    lock.releaseLock();
  }
}



const PASSWORD_HASH_VERSION = "SHA256I";
const PASSWORD_HASH_ITERATIONS = 5000;
const PASSWORD_MAX_LENGTH = 128;
const PASSWORD_PEPPER_PROPERTY = "ADMIN_PASSWORD_PEPPER";
const PASSWORD_VERIFIER_PREFIX = "ADMIN_PASSWORD_VERIFIER_";

function bytesToHex(bytes) {
  return bytes
    .map(byte => ("0" + ((byte + 256) % 256).toString(16)).slice(-2))
    .join("");
}

function sha256Hex(value) {
  return bytesToHex(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(value),
      Utilities.Charset.UTF_8
    )
  );
}

function constantTimeEqual(left, right) {
  left = String(left || "");
  right = String(right || "");

  if (left.length !== right.length) {
    return false;
  }

  let difference = 0;

  for (let i = 0; i < left.length; i++) {
    difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
  }

  return difference === 0;
}

function deriveIterativePasswordHash(password, salt, iterations) {
  let hash = sha256Hex(salt + ":" + password);

  for (let i = 1; i < iterations; i++) {
    hash = sha256Hex(hash + ":" + salt + ":" + password);
  }

  return hash;
}

function createPasswordHash(password) {
  const salt = Utilities.getUuid().replace(/-/g, "");
  const hash = deriveIterativePasswordHash(
    String(password),
    salt,
    PASSWORD_HASH_ITERATIONS
  );

  return [
    PASSWORD_HASH_VERSION,
    PASSWORD_HASH_ITERATIONS,
    salt,
    hash
  ].join("$");
}

function getPasswordPepper() {
  const properties = PropertiesService.getScriptProperties();
  let pepper = String(
    properties.getProperty(PASSWORD_PEPPER_PROPERTY) || ""
  );

  if (pepper) {
    return pepper;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    pepper = String(
      properties.getProperty(PASSWORD_PEPPER_PROPERTY) || ""
    );

    if (!pepper) {
      pepper = Utilities.getUuid().replace(/-/g, "") +
        Utilities.getUuid().replace(/-/g, "");
      properties.setProperty(PASSWORD_PEPPER_PROPERTY, pepper);
    }

    return pepper;
  } finally {
    lock.releaseLock();
  }
}

function getPasswordVerifierKey(username) {
  const normalizedUsername =
    String(username || "").trim().toLowerCase();

  return PASSWORD_VERIFIER_PREFIX + sha256Hex(normalizedUsername);
}

function createCheapPasswordVerifier(username, password) {
  const normalizedUsername =
    String(username || "").trim().toLowerCase();

  return sha256Hex(
    normalizedUsername + ":" +
    String(password) + ":" +
    getPasswordPepper()
  );
}

function getCheapPasswordVerifier(username) {
  return String(
    PropertiesService
      .getScriptProperties()
      .getProperty(getPasswordVerifierKey(username)) || ""
  );
}

function setCheapPasswordVerifier(username, verifier) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      getPasswordVerifierKey(username),
      verifier
    );
}

function clearCheapPasswordVerifier(username) {
  PropertiesService
    .getScriptProperties()
    .deleteProperty(getPasswordVerifierKey(username));
}

function clearLegacyLoginAttempts(username) {
  const normalizedUsername =
    String(username || "").trim().toLowerCase();

  PropertiesService
    .getScriptProperties()
    .deleteProperty("LOGIN_ATTEMPTS_" + sha256Hex(normalizedUsername));
}

function isLegacyPasswordHash(storedHash) {
  return /^[a-f0-9]{64}$/i.test(String(storedHash || "").trim());
}

function verifyPassword(password, storedHash) {
  const normalizedHash = String(storedHash || "").trim();

  if (isLegacyPasswordHash(normalizedHash)) {
    return constantTimeEqual(
      sha256Hex(String(password)),
      normalizedHash.toLowerCase()
    );
  }

  const parts = normalizedHash.split("$");
  if (
    parts.length !== 4 ||
    parts[0] !== PASSWORD_HASH_VERSION
  ) {
    return false;
  }

  const iterations = Number(parts[1]);
  const salt = parts[2];
  const expectedHash = String(parts[3] || "").toLowerCase();

  if (
    !Number.isInteger(iterations) ||
    iterations < 1 ||
    iterations > PASSWORD_HASH_ITERATIONS ||
    !/^[a-f0-9]{16,128}$/i.test(salt) ||
    !/^[a-f0-9]{64}$/.test(expectedHash)
  ) {
    return false;
  }

  return constantTimeEqual(
    deriveIterativePasswordHash(String(password), salt, iterations),
    expectedHash
  );
}

function adminLogin(username, password) {
  const ss = getSS();
  const adminSheet   = ss.getSheetByName("admins");
  const sessionSheet = ss.getSheetByName("sessions");

  if (!adminSheet || !sessionSheet) {
    return {
      success: false,
      message: "System not ready"
    };
  }

  // normalize input
  username = String(username || "").trim();
  password = String(password || "").trim();

  if (!username || !password || password.length > PASSWORD_MAX_LENGTH) {
    return {
      success: false,
      message: "Username หรือ Password ไม่ถูกต้อง"
    };
  }

  const lastRow = adminSheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: false,
      message: "No admin configured"
    };
  }

  // load admins (username, password_hash)
  const admins = adminSheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .map(r => [
      String(r[0]).trim(), // username
      String(r[1]).trim()  // password_hash
    ]);

  const adminRecord = admins.find(r => r[0] === username);

  if (!adminRecord) {
    return {
      success: false,
      message: "Username หรือ Password ไม่ถูกต้อง"
    };
  }

  const cheapVerifier = getCheapPasswordVerifier(username);

  if (
    cheapVerifier &&
    !constantTimeEqual(
      createCheapPasswordVerifier(username, password),
      cheapVerifier
    )
  ) {
    return {
      success: false,
      message: "Username หรือ Password ไม่ถูกต้อง"
    };
  }

  const found = verifyPassword(password, adminRecord[1])
    ? adminRecord
    : null;

  if (!found) {
    return {
      success: false,
      message: "Username หรือ Password ไม่ถูกต้อง"
    };
  }

  const verifiedCheapPassword = createCheapPasswordVerifier(
    username,
    password
  );
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const currentAdminLastRow = adminSheet.getLastRow();
    if (currentAdminLastRow < 2) {
      return {
        success: false,
        message: "No admin configured"
      };
    }

    const currentAdmins = adminSheet
      .getRange(2, 1, currentAdminLastRow - 1, 2)
      .getValues();
    const currentAdminIndex = currentAdmins.findIndex(
      row =>
        String(row[0]).trim() === username &&
        constantTimeEqual(
          String(row[1] || "").trim(),
          String(found[1] || "").trim()
        )
    );

    if (currentAdminIndex === -1) {
      return {
        success: false,
        message: "Username หรือ Password ไม่ถูกต้อง"
      };
    }

    const storedHash = String(currentAdmins[currentAdminIndex][1] || "").trim();
    if (isLegacyPasswordHash(storedHash)) {
      adminSheet
        .getRange(currentAdminIndex + 2, 2)
        .setValue(createPasswordHash(password));
    }

    setCheapPasswordVerifier(username, verifiedCheapPassword);
    clearLegacyLoginAttempts(username);

    const now = new Date();
    const sessionRows = sessionSheet.getDataRange().getValues();

    for (let i = sessionRows.length - 1; i >= 1; i--) {
      const sessionUsername = String(sessionRows[i][1] || "").trim();
      const expiredAt = sessionRows[i][2];

      if (
        sessionUsername === username ||
        !expiredAt ||
        new Date(expiredAt) <= now
      ) {
        sessionSheet.deleteRow(i + 1);
      }
    }

    const token = Utilities.getUuid();
    const expiredAt = new Date(Date.now() + 12 * 60 * 60 * 1000);

    sessionSheet.appendRow([
      token,
      username,
      expiredAt
    ]);

    return {
      success: true,
      token,
      username,
      expiredAt
    };
  } finally {
    lock.releaseLock();
  }
}

function adminLogout(token) {
  token = String(token || "").trim();

  if (!token) {
    throw new Error("Missing token");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sessionSheet = getSS().getSheetByName("sessions");

    if (!sessionSheet) {
      throw new Error("Session system not ready");
    }

    const lastRow = sessionSheet.getLastRow();

    if (lastRow < 2) {
      throw new Error("Invalid token");
    }

    const tokens = sessionSheet
      .getRange(2, 1, lastRow - 1, 1)
      .getValues();

    const sessionIndex = tokens.findIndex(
      row => String(row[0]).trim() === token
    );

    if (sessionIndex === -1) {
      throw new Error("Invalid token");
    }

    sessionSheet.deleteRow(sessionIndex + 2);

    return {
      success: true
    };
  } finally {
    lock.releaseLock();
  }
}

function changePassword(token, currentPassword, newPassword) {
  const auth = requireAuth(token);
  const username = auth.username;

  currentPassword = String(currentPassword || "").trim();
  newPassword = String(newPassword || "").trim();

  if (!currentPassword || !newPassword) {
    throw new Error("Missing password data");
  }

  if (
    currentPassword.length > PASSWORD_MAX_LENGTH ||
    newPassword.length > PASSWORD_MAX_LENGTH
  ) {
    throw new Error("Password is too long");
  }

  if (newPassword.length < 8) {
    throw new Error("New password must be at least 8 characters");
  }

  const newHash = createPasswordHash(newPassword);
  const newCheapPasswordVerifier = createCheapPasswordVerifier(
    username,
    newPassword
  );

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const lockedAuth = requireAuth(token);
    if (String(lockedAuth.username) !== String(username)) {
      throw new Error("Invalid token");
    }

    const ss = getSS();
    const adminSheet = ss.getSheetByName("admins");
    const sessionSheet = ss.getSheetByName("sessions");

    if (!adminSheet || !sessionSheet) {
      throw new Error("System not ready");
    }

    const lastRow = adminSheet.getLastRow();
    if (lastRow < 2) {
      throw new Error("No admin configured");
    }

    const rows = adminSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const rowIndex = rows.findIndex(r =>
      String(r[0]).trim() === username &&
      verifyPassword(currentPassword, r[1])
    );

    if (rowIndex === -1) {
      throw new Error("Current password is incorrect");
    }

    clearCheapPasswordVerifier(username);
    adminSheet.getRange(rowIndex + 2, 2).setValue(newHash);
    setCheapPasswordVerifier(username, newCheapPasswordVerifier);

    const sessionRows = sessionSheet.getDataRange().getValues();

    for (let i = sessionRows.length - 1; i >= 1; i--) {
      if (String(sessionRows[i][1]).trim() === username) {
        sessionSheet.deleteRow(i + 1);
      }
    }

    return {
      success: true,
      username
    };
  } finally {
    lock.releaseLock();
  }
}

function requireAuth(token) {
  if (!token) {
    throw new Error("Missing token");
  }

  const ss = getSS();
  const sheet = ss.getSheetByName("sessions");

  if (!sheet) {
    throw new Error("Session system not ready");
  }

  const lastRow = sheet.getLastRow();

  // 🔒 GUARD: ไม่มี session จริง (มีแต่ header หรือว่าง)
  if (lastRow < 2) {
    throw new Error("Invalid token");
  }

  const rows = sheet
    .getRange(2, 1, lastRow - 1, 3)
    .getValues();

  const now = new Date();

  for (const [savedToken, username, expiredAt] of rows) {
    if (String(savedToken) === String(token)) {

      if (!expiredAt || new Date(expiredAt) <= now) {
        throw new Error("Session expired");
      }

      return {
        username,
        token,
        role: "admin", // 🔒 เตรียมไว้
        expiredAt: new Date(expiredAt)
      };
    }
  }

  throw new Error("Invalid token");
}




function cleanupExpiredSessions() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getSS().getSheetByName("sessions");

    if (!sheet) {
      Logger.log("cleanupExpiredSessions: sessions sheet not found");
      return;
    }

    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      Logger.log("cleanupExpiredSessions: no session rows to clean");
      return;
    }

    const now = new Date();
    let removed = 0;

    for (let i = rows.length - 1; i >= 1; i--) {
      const expiredAt = rows[i][2];

      if (!expiredAt || new Date(expiredAt) <= now) {
        sheet.deleteRow(i + 1);
        removed++;
      }
    }

    Logger.log(
      `cleanupExpiredSessions: removed ${removed} expired session(s)`
    );
  } finally {
    lock.releaseLock();
  }
}

function getOrders() {
  // 🔒 FIX: ใช้ Spreadsheet เดียวกับทั้งระบบ
  const ss = getSS();
  const sheet = ss.getSheetByName("Orders");

  if (!sheet) {
    return [];
  }

  ensureOrdersExtendedHeaders_(ss);

  const rows = sheet.getDataRange().getValues();

  // มีแต่ header
  if (rows.length < 2) {
    return [];
  }

  const headers = rows.shift();

  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = r[i];
    });

    // 🔧 parse items จาก string → array
    if (typeof obj.items === "string") {
      try {
        obj.items = JSON.parse(obj.items);
      } catch (e) {
        obj.items = [];
      }
    }

    return obj;
  });
}


function normalizePoNumberForLookup_(value) {
  return String(value || "").trim().toLowerCase();
}

function checkDuplicatePo(params) {
  const poNumber = String(params && params.poNumber || "").trim();
  const normalizedPo = normalizePoNumberForLookup_(poNumber);
  const data = {
    duplicate: false,
    poNumber,
    matches: []
  };

  if (!normalizedPo) {
    return {
      success: true,
      data
    };
  }

  const ss = getSS();
  const sheet = ss.getSheetByName("Orders");
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      success: true,
      data
    };
  }

  const values = sheet.getDataRange().getValues();
  const headers = values.shift().map(header => String(header || "").trim());
  const col = {
    orderId: headers.indexOf("orderId"),
    status: headers.indexOf("status"),
    poNumber: headers.indexOf("poNumber"),
    source: headers.indexOf("source"),
    partnerRequestId: headers.indexOf("partnerRequestId")
  };

  if (col.poNumber < 0) {
    return {
      success: true,
      data
    };
  }

  data.matches = values
    .filter(row => normalizePoNumberForLookup_(row[col.poNumber]) === normalizedPo)
    .map(row => ({
      orderId: col.orderId >= 0 ? String(row[col.orderId] || "") : "",
      status: col.status >= 0 ? String(row[col.status] || "") : "",
      source: col.source >= 0
        ? String(row[col.source] || "viewer")
        : "viewer",
      partnerRequestId: col.partnerRequestId >= 0
        ? String(row[col.partnerRequestId] || "")
        : ""
    }));
  data.duplicate = data.matches.length > 0;

  return {
    success: true,
    data
  };
}


function addProduct(e, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const by = auth.username;

  /* ================= READ PARAM ================= */
  const id     = String(e.parameter.productId || "").trim();
  const name   = String(e.parameter.name || "").trim();
  const price  = Number(e.parameter.price);
  const rawCostPrice = String(e.parameter.costPrice ?? "").trim();
  const costPrice = rawCostPrice === "" ? null : Number(rawCostPrice);
  const stock  = Number(e.parameter.stock);
  const active = true; // สินค้าใหม่เปิดขายเสมอ
  const image  = String(e.parameter.image || "").trim();
  let   status = String(e.parameter.status || "ready").trim();
  const note   = String(e.parameter.note || "").trim();
  const variantMetadata = parseProductVariantMetadata_(e.parameter);

  /* ================= VALIDATE ================= */
  if (!id || !name ||
     !Number.isInteger(price) ||
     costPrice === null ||
     !Number.isInteger(costPrice) ||
     !Number.isInteger(stock)) {
    throw new Error("Invalid product data");
  }

  if (price < 0 || costPrice < 0 || stock < 0) {
    throw new Error("Price, cost price and stock must be >= 0");
  }

  // 🔒 auto status: stock = 0 → out
  if (stock === 0) {
    status = "out";
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // 🔒 ใช้ Spreadsheet เดียวกับทั้งระบบ
    const ss = getSS();

    const sh = ss.getSheetByName("Products");
    if (!sh) throw new Error("Products sheet not found");
    const productHeaders = ensureProductsVariantHeaders_(ss).headers;

    // 🆕 LOG SHEET
    const logSheet = ss.getSheetByName("stock_logs");
    if (!logSheet) throw new Error("stock_logs sheet not found");

    // 🔍 กัน SKU ซ้ำภายใน lock
    const rows = sh.getDataRange().getValues();
    const exists = rows.slice(1).some(r => String(r[0]).trim() === id);
    if (exists) {
      throw new Error("SKU already exists");
    }

    const createdRowNumber = sh.getLastRow() + 1;

    /* ================= ADD PRODUCT ================= */
    sh.appendRow([
      id,          // A productId
      name,        // B name
      price,       // C price
      stock,       // D stock
      image,       // E image
      active,      // F active
      status,      // G status  ✅ เพิ่มใหม่
      new Date(),  // H createdAt
      by,          // I createdBy
      note,        // J note
      "",          // K detailsText
      "",          // L compareImages
      costPrice    // M costPrice
    ]);

    try {
      writeProductVariantMetadata_(
        sh,
        createdRowNumber,
        variantMetadata,
        productHeaders
      );

      /* ================= LOG CREATE ================= */
      logSheet.appendRow([
        "LOG-" + Date.now(), // logId
        id,                  // productId
        "CREATE",            // type
        stock,               // qty
        0,                   // before
        stock,               // after
        by,                  // by
        "",                  // orderId
        "CREATE_PRODUCT",    // reason
        new Date()           // timestamp
      ]);
    } catch (err) {
      try {
        const createdProductId = String(
          sh.getRange(createdRowNumber, 1).getValue()
        ).trim();

        if (createdProductId !== id) {
          throw new Error("Created product row no longer matches");
        }

        sh.deleteRow(createdRowNumber);
      } catch (rollbackErr) {
        throw new Error(
          String(err.message || err) +
          " | Product create rollback failed: " +
          String(rollbackErr.message || rollbackErr)
        );
      }

      throw err;
    }

    return {
      success: true,
      createdBy: by
    };
  } finally {
    lock.releaseLock();
  }
}

function bulkAddProducts(e, auth) {
  if (!auth || !auth.username) {
    return bulkRuntimeFailure("Unauthorized");
  }

  const by = auth.username;
  const lock = LockService.getScriptLock();

  try {
    const parsedItems = parseBulkProductItems(e.parameter.items);
    if (parsedItems.errors.length) {
      return {
        success: false,
        error: "Bulk import validation failed",
        data: {
          imported: 0,
          errors: parsedItems.errors
        }
      };
    }

    const items = parsedItems.items;

    lock.waitLock(30000);

    const ss = getSS();
    const productSheet = ss.getSheetByName("Products");
    if (!productSheet) {
      throw new Error("Products sheet not found");
    }
    const productHeaders = ensureProductsVariantHeaders_(ss).headers;

    const logSheet = ss.getSheetByName("stock_logs");
    if (!logSheet) {
      throw new Error("stock_logs sheet not found");
    }

    const productRows = productSheet.getDataRange().getValues();
    const existingSkus = new Set(
      productRows
        .slice(1)
        .map(row => String(row[0] || "").trim().toUpperCase())
        .filter(Boolean)
    );

    const validation = validateBulkProductItems(items, existingSkus);
    if (validation.errors.length) {
      return {
        success: false,
        error: "Bulk import validation failed",
        data: {
          imported: 0,
          errors: validation.errors
        }
      };
    }

    const createdProductIds = [];
    const createdLogIds = [];
    const createdAt = new Date();

    try {
      validation.products.forEach(product => {
        productSheet.appendRow([
          product.productId,
          product.name,
          product.price,
          product.stock,
          product.image,
          product.active,
          product.status,
          createdAt,
          by,
          product.note,
          product.detailsText,
          product.compareImages,
          product.costPrice
        ]);
        writeProductVariantMetadata_(
          productSheet,
          productSheet.getLastRow(),
          product.variantMetadata,
          productHeaders
        );
        createdProductIds.push(product.productId);
      });

      validation.products.forEach(product => {
        const logId = "LOG-" + Utilities.getUuid();
        logSheet.appendRow([
          logId,
          product.productId,
          "CREATE",
          product.stock,
          0,
          product.stock,
          by,
          "",
          "BULK_CREATE_PRODUCT",
          new Date()
        ]);
        createdLogIds.push(logId);
      });
    } catch (err) {
      try {
        rollbackBulkProductImport(
          productSheet,
          logSheet,
          createdProductIds,
          createdLogIds
        );
      } catch (rollbackErr) {
        throw new Error(
          String(err.message || err) +
          " | Bulk import rollback failed: " +
          String(rollbackErr.message || rollbackErr)
        );
      }
      throw err;
    }

    return {
      success: true,
      imported: validation.products.length,
      failed: 0,
      errors: []
    };
  } catch (err) {
    Logger.log("bulkAddProducts error: " + err);
    return bulkRuntimeFailure(
      err && err.message ? err.message : "Bulk import failed"
    );
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {
      // Ignore release errors when lock acquisition failed before waitLock.
    }
  }
}

function bulkUpdateProducts(e, auth) {
  if (!auth || !auth.username) {
    return bulkUpdateRuntimeFailure("Unauthorized");
  }

  const lock = LockService.getScriptLock();

  try {
    const parsedItems = parseBulkUpdateProductItems(e.parameter.items);
    if (parsedItems.errors.length) {
      return bulkUpdateValidationFailure(parsedItems.errors);
    }

    lock.waitLock(30000);

    const ss = getSS();
    const productSheet = ss.getSheetByName("Products");
    if (!productSheet) {
      throw new Error("Products sheet not found");
    }

    const variantHeaderResult = ensureProductsVariantHeaders_(ss);
    const productHeaders = variantHeaderResult.headers;
    const productRows = productSheet.getDataRange().getValues();
    const validation = validateBulkUpdateProductItems(
      parsedItems.items,
      productRows,
      productHeaders
    );

    if (validation.errors.length) {
      return bulkUpdateValidationFailure(validation.errors);
    }

    const changedRows = [];

    try {
      validation.updates.forEach(update => {
        const rowNumber = update.rowIndex + 1;
        const originalRow = productRows[update.rowIndex].slice();
        const nextRow = originalRow.slice();

        if (update.hasOwnProperty("price")) {
          nextRow[2] = update.price;
        }
        if (update.hasOwnProperty("active")) {
          nextRow[5] = update.active;
        }
        if (update.hasOwnProperty("status")) {
          nextRow[6] = update.status;
        }
        if (update.hasOwnProperty("costPrice")) {
          nextRow[12] = update.costPrice;
        }
        if (update.variantMetadata) {
          const variantColumns = getProductVariantColumnMap_(productHeaders);
          PRODUCT_OPTIONAL_VARIANT_HEADERS.forEach(header => {
            if (!Object.prototype.hasOwnProperty.call(update.variantMetadata, header)) return;
            const columnIndex = variantColumns[header];
            if (columnIndex === undefined || columnIndex < 0) return;
            nextRow[columnIndex] = update.variantMetadata[header] === undefined ||
              update.variantMetadata[header] === null
                ? ""
                : update.variantMetadata[header];
          });
        }

        productSheet
          .getRange(rowNumber, 1, 1, nextRow.length)
          .setValues([nextRow]);

        changedRows.push({
          rowNumber,
          values: originalRow
        });
      });
    } catch (err) {
      try {
        rollbackBulkProductUpdates(productSheet, changedRows);
      } catch (rollbackErr) {
        throw new Error(
          String(err.message || err) +
          " | Bulk update rollback failed: " +
          String(rollbackErr.message || rollbackErr)
        );
      }
      throw err;
    }

    return {
      success: true,
      updated: validation.updates.length,
      failed: 0,
      errors: []
    };
  } catch (err) {
    Logger.log("bulkUpdateProducts error: " + err);
    return bulkUpdateRuntimeFailure(
      err && err.message ? err.message : "Bulk update failed"
    );
  } finally {
    try {
      lock.releaseLock();
    } catch (releaseErr) {
      // Ignore release errors when lock acquisition failed before waitLock.
    }
  }
}

function parseBulkUpdateProductItems(rawItems) {
  let items;
  try {
    items = JSON.parse(String(rawItems || "[]"));
  } catch (err) {
    return {
      items: [],
      errors: [
        bulkProductError(0, "", "items", "Items must be valid JSON")
      ]
    };
  }

  if (!Array.isArray(items)) {
    return {
      items: [],
      errors: [
        bulkProductError(0, "", "items", "Items must be an array")
      ]
    };
  }

  return {
    items,
    errors: []
  };
}

function validateBulkUpdateProductItems(items, productRows, productHeaders) {
  const errors = [];
  const updates = [];
  const batchSkus = new Set();
  const allowedFields = [
    "productId",
    "price",
    "costPrice",
    "status",
    "active"
  ].concat(PRODUCT_OPTIONAL_VARIANT_HEADERS);
  const updateFields = [
    "price",
    "costPrice",
    "status",
    "active"
  ].concat(PRODUCT_OPTIONAL_VARIANT_HEADERS);
  const allowedStatuses = ["ready", "ready_plus", "shipping", "warehouse", "out"];
  const productMap = new Map();
  const variantHeaders = Array.isArray(productHeaders) && productHeaders.length
    ? productHeaders
    : [];
  const variantColumns = getProductVariantColumnMap_(variantHeaders);

  productRows.slice(1).forEach((row, index) => {
    const productId = String(row[0] || "").trim().toUpperCase();
    if (productId) {
      productMap.set(productId, {
        rowIndex: index + 1,
        row
      });
    }
  });

  if (items.length === 0) {
    errors.push(bulkProductError(1, "", "items", "At least one product update is required"));
  }

  if (items.length > 100) {
    errors.push(bulkProductError(0, "", "items", "Batch size limit is 100 rows"));
    return {
      updates: [],
      errors
    };
  }

  items.forEach((item, index) => {
    const rowNumber = index + 1;

    if (!item || typeof item !== "object" || Array.isArray(item)) {
      errors.push(bulkProductError(rowNumber, "", "item", "Item must be an object"));
      return;
    }

    const keys = Object.keys(item);
    const productId = String(item.productId || "").trim().toUpperCase();
    const update = {
      productId
    };
    let hasAllowedUpdate = false;

    keys.forEach(key => {
      if (key === "stock") {
        errors.push(bulkProductError(rowNumber, productId, key, "Stock cannot be updated by bulkUpdateProducts"));
      } else if (allowedFields.indexOf(key) === -1) {
        errors.push(bulkProductError(rowNumber, productId, key, "Unknown field is not allowed"));
      }
    });

    if (!productId) {
      errors.push(bulkProductError(rowNumber, productId, "productId", "SKU is required"));
    } else if (batchSkus.has(productId)) {
      errors.push(bulkProductError(rowNumber, productId, "productId", "Duplicate SKU inside update batch"));
    } else if (!productMap.has(productId)) {
      errors.push(bulkProductError(rowNumber, productId, "productId", "Product not found"));
    }

    if (productId) {
      batchSkus.add(productId);
    }

    updateFields.forEach(field => {
      if (Object.prototype.hasOwnProperty.call(item, field)) {
        hasAllowedUpdate = true;
      }
    });

    if (!hasAllowedUpdate) {
      errors.push(bulkProductError(rowNumber, productId, "items", "At least one allowed field is required"));
    }

    if (Object.prototype.hasOwnProperty.call(item, "price")) {
      const price = Number(item.price);
      if (isBlankBulkValue(item.price) || !Number.isInteger(price) || price < 0) {
        errors.push(bulkProductError(rowNumber, productId, "price", "Price must be an integer >= 0"));
      } else {
        update.price = price;
      }
    }

    if (Object.prototype.hasOwnProperty.call(item, "costPrice")) {
      const costPrice = Number(item.costPrice);
      if (isBlankBulkValue(item.costPrice) || !Number.isInteger(costPrice) || costPrice < 0) {
        errors.push(bulkProductError(rowNumber, productId, "costPrice", "Cost price must be an integer >= 0"));
      } else {
        update.costPrice = costPrice;
      }
    }

    if (Object.prototype.hasOwnProperty.call(item, "status")) {
      const status = String(item.status || "").trim();
      if (!status || allowedStatuses.indexOf(status) === -1) {
        errors.push(bulkProductError(rowNumber, productId, "status", "Invalid product status"));
      } else {
        update.status = status;
      }
    }

    if (Object.prototype.hasOwnProperty.call(item, "active")) {
      const activeResult = parseBulkUpdateProductActive(item.active);
      if (!activeResult.valid) {
        errors.push(bulkProductError(rowNumber, productId, "active", "Active must be true/false or 1/0"));
      } else {
        update.active = activeResult.value;
      }
    }

    const hasVariantMetadata = PRODUCT_OPTIONAL_VARIANT_HEADERS.some(field =>
      Object.prototype.hasOwnProperty.call(item, field)
    );
    if (hasVariantMetadata) {
      try {
        const variantMetadata = parseProductVariantMetadata_(item);
        const selectedVariantMetadata = {};
        PRODUCT_OPTIONAL_VARIANT_HEADERS.forEach(field => {
          if (Object.prototype.hasOwnProperty.call(item, field)) {
            if (variantColumns[field] === undefined || variantColumns[field] < 0) {
              errors.push(bulkProductError(rowNumber, productId, field, "Variant metadata column is missing"));
            }
            selectedVariantMetadata[field] = variantMetadata[field];
          }
        });
        update.variantMetadata = selectedVariantMetadata;
      } catch (err) {
        errors.push(
          bulkProductError(
            rowNumber,
            productId,
            "variantSortOrder",
            err && err.message ? err.message : "Variant sort order must be an integer"
          )
        );
      }
    }

    if (productMap.has(productId)) {
      update.rowIndex = productMap.get(productId).rowIndex;
    }

    updates.push(update);
  });

  return {
    updates,
    errors
  };
}

function parseBulkUpdateProductActive(value) {
  if (value === true || value === false) {
    return {
      valid: true,
      value
    };
  }

  const normalized = String(value).trim().toLowerCase();
  if (normalized === "true" || normalized === "1") {
    return {
      valid: true,
      value: true
    };
  }

  if (normalized === "false" || normalized === "0") {
    return {
      valid: true,
      value: false
    };
  }

  return {
    valid: false,
    value: true
  };
}

function bulkUpdateValidationFailure(errors) {
  return {
    success: false,
    error: "Bulk update validation failed",
    data: {
      updated: 0,
      errors
    }
  };
}

function bulkUpdateRuntimeFailure(message) {
  return {
    success: false,
    error: message || "Bulk update failed",
    data: {
      updated: 0,
      errors: []
    }
  };
}

function rollbackBulkProductUpdates(productSheet, changedRows) {
  if (!changedRows || !changedRows.length) return;

  for (let i = changedRows.length - 1; i >= 0; i--) {
    const changed = changedRows[i];
    productSheet
      .getRange(changed.rowNumber, 1, 1, changed.values.length)
      .setValues([changed.values]);
  }
}

function parseBulkProductItems(rawItems) {
  let items;
  try {
    items = JSON.parse(String(rawItems || "[]"));
  } catch (err) {
    return {
      items: [],
      errors: [
        bulkProductError(0, "", "items", "Items must be valid JSON")
      ]
    };
  }

  if (!Array.isArray(items)) {
    return {
      items: [],
      errors: [
        bulkProductError(0, "", "items", "Items must be an array")
      ]
    };
  }

  return {
    items,
    errors: []
  };
}

function validateBulkProductItems(items, existingSkus) {
  const errors = [];
  const products = [];
  const batchSkus = new Set();

  if (items.length === 0) {
    errors.push(bulkProductError(1, "", "items", "At least one product is required"));
  }

  if (items.length > 100) {
    errors.push(bulkProductError(0, "", "items", "Batch size limit is 100 rows"));
  }

  items.slice(0, 100).forEach((item, index) => {
    const rowNumber = index + 1;
    const productId = String(item && item.productId || "").trim().toUpperCase();
    const name = String(item && item.name || "").trim();
    const rawPrice = item ? item.price : "";
    const rawCostPrice = item ? item.costPrice : "";
    const rawStock = item ? item.stock : "";
    const price = Number(rawPrice);
    const costPrice = Number(rawCostPrice);
    const stock = Number(rawStock);
    const rawStatus = String(item && item.status || "").trim();
    const image = String(item && item.image || "").trim();
    const note = String(item && item.note || "").trim();
    const detailsText = String(item && item.detailsText || "").trim();
    const compareImages = String(item && item.compareImages || "").trim();
    let variantMetadata;
    const activeResult = parseBulkProductActive(item ? item.active : undefined);
    const allowedStatuses = ["ready", "ready_plus", "shipping", "warehouse", "out"];
    let status = rawStatus;

    try {
      variantMetadata = parseProductVariantMetadata_(item || {});
    } catch (err) {
      variantMetadata = parseProductVariantMetadata_({});
      errors.push(
        bulkProductError(
          rowNumber,
          productId,
          "variantSortOrder",
          err && err.message ? err.message : "Variant sort order must be an integer"
        )
      );
    }

    if (!productId) {
      errors.push(bulkProductError(rowNumber, productId, "productId", "SKU is required"));
    } else if (batchSkus.has(productId)) {
      errors.push(bulkProductError(rowNumber, productId, "productId", "Duplicate SKU inside import batch"));
    } else if (existingSkus.has(productId)) {
      errors.push(bulkProductError(rowNumber, productId, "productId", "SKU already exists"));
    }

    if (productId) {
      batchSkus.add(productId);
    }

    if (!name) {
      errors.push(bulkProductError(rowNumber, productId, "name", "Product name is required"));
    }

    if (isBlankBulkValue(rawPrice) || !Number.isInteger(price) || price < 0) {
      errors.push(bulkProductError(rowNumber, productId, "price", "Price must be an integer >= 0"));
    }

    if (isBlankBulkValue(rawCostPrice) || !Number.isInteger(costPrice) || costPrice < 0) {
      errors.push(bulkProductError(rowNumber, productId, "costPrice", "Cost price must be an integer >= 0"));
    }

    if (isBlankBulkValue(rawStock) || !Number.isInteger(stock) || stock < 0) {
      errors.push(bulkProductError(rowNumber, productId, "stock", "Stock must be an integer >= 0"));
    }

    if (!status) {
      errors.push(bulkProductError(rowNumber, productId, "status", "Status is required"));
    } else if (allowedStatuses.indexOf(status) === -1) {
      errors.push(bulkProductError(rowNumber, productId, "status", "Invalid product status"));
    }

    if (!activeResult.valid) {
      errors.push(bulkProductError(rowNumber, productId, "active", "Active must be true/false or 1/0"));
    }

    if (image && !/^https?:\/\//i.test(image)) {
      errors.push(bulkProductError(rowNumber, productId, "image", "Image must be an http/https URL"));
    }

    if (compareImages && !compareImages.split(",").every(isHttpUrlText)) {
      errors.push(bulkProductError(rowNumber, productId, "compareImages", "Compare images must be comma-separated http/https URLs"));
    }

    if (note.length > 2000) {
      errors.push(bulkProductError(rowNumber, productId, "note", "Note must be 2000 characters or fewer"));
    }

    if (detailsText.length > 5000) {
      errors.push(bulkProductError(rowNumber, productId, "detailsText", "Details text must be 5000 characters or fewer"));
    }

    if (Number.isInteger(stock) && stock === 0) {
      status = "out";
    }

    products.push({
      productId,
      name,
      price,
      costPrice,
      stock,
      image,
      active: activeResult.value,
      status,
      note,
      detailsText,
      compareImages,
      variantMetadata
    });
  });

  return {
    products,
    errors
  };
}

function parseBulkProductActive(value) {
  if (value === undefined || value === null || value === "") {
    return {
      valid: true,
      value: true
    };
  }

  if (value === true || value === false) {
    return {
      valid: true,
      value
    };
  }

  const normalized = String(value).trim().toLowerCase();
  if (normalized === "true" || normalized === "1") {
    return {
      valid: true,
      value: true
    };
  }

  if (normalized === "false" || normalized === "0") {
    return {
      valid: true,
      value: false
    };
  }

  return {
    valid: false,
    value: true
  };
}

function isBlankBulkValue(value) {
  return value === undefined ||
    value === null ||
    String(value).trim() === "";
}

function isHttpUrlText(value) {
  const text = String(value || "").trim();
  return !text || /^https?:\/\//i.test(text);
}

function bulkProductError(row, productId, field, message) {
  return {
    row,
    productId,
    field,
    message
  };
}

function bulkRuntimeFailure(message) {
  return {
    success: false,
    error: message || "Bulk import failed",
    data: {
      imported: 0,
      errors: []
    }
  };
}

function rollbackBulkProductImport(productSheet, logSheet, productIds, logIds) {
  if (logIds && logIds.length) {
    const logIdSet = new Set(logIds);
    const logRows = logSheet.getDataRange().getValues();
    for (let i = logRows.length - 1; i >= 1; i--) {
      if (logIdSet.has(String(logRows[i][0] || "").trim())) {
        logSheet.deleteRow(i + 1);
      }
    }
  }

  if (productIds && productIds.length) {
    const productIdSet = new Set(
      productIds.map(id => String(id || "").trim().toUpperCase())
    );
    const productRows = productSheet.getDataRange().getValues();
    for (let i = productRows.length - 1; i >= 1; i--) {
      const productId = String(productRows[i][0] || "").trim().toUpperCase();
      if (productIdSet.has(productId)) {
        productSheet.deleteRow(i + 1);
      }
    }
  }
}

function cleanupUnusedImages({ by, dryRun }) {
  const ss = getSS();
  const productSheet = ss.getSheetByName("Products");
  if (!productSheet) {
    throw new Error("Products sheet not found");
  }

  /* ================= COLLECT USED IMAGE IDS ================= */
  const rows = productSheet.getDataRange().getValues();
  rows.shift(); // header

  const usedFileIds = new Set();

  rows.forEach(r => {
    const imageUrl = String(r[4] || "").trim();
    if (!imageUrl) return;

    // extract fileId from lh3 url
    const match = imageUrl.match(/\/d\/([^=]+)/);
    if (match && match[1]) {
      usedFileIds.add(match[1]);
    }
  });

  /* ================= SCAN DRIVE FOLDER ================= */
  const FOLDER_ID = "1un_A6DFFnknmEjx7LgACRKT7l8AgBBdK";
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  let scanned = 0;
  let orphan = [];
  let removed = 0;

  while (files.hasNext()) {
    const file = files.next();
    scanned++;

    const fileId = file.getId();
    if (!usedFileIds.has(fileId)) {
      orphan.push({
        fileId,
        name: file.getName()
      });

      if (!dryRun) {
        file.setTrashed(true);
        removed++;
      }
    }
  }

  /* ================= LOG ================= */
  Logger.log(
    `[GC] by=${by} dryRun=${dryRun} scanned=${scanned} orphan=${orphan.length}`
  );

  return {
    dryRun,
    scanned,
    orphanCount: orphan.length,
    removed,
    orphanFiles: orphan
  };
}

function generateOrderPDF(token, orderId) {
  const auth = requireAuth(token);
  const by = auth.username;

  if (!orderId) {
    throw new Error("Missing orderId");
  }

  const ss = getSS();
  const sheet = ss.getSheetByName("Orders");
  if (!sheet) {
    throw new Error("Orders sheet not found");
  }

  const rows = sheet.getDataRange().getValues();
  rows.shift(); // remove header

  const idx = rows.findIndex(r => r[0] === orderId);
  if (idx === -1) {
    throw new Error("Order not found");
  }

  const row = rows[idx];

  let items = [];
  try {
    items = JSON.parse(row[1] || "[]");
    if (!Array.isArray(items)) items = [];
  } catch (err) {
    items = [];
  }
  const total = Number(row[2]) || 0;
  const status = String(row[3] || "");
  const createdAtRaw = row[4] ? new Date(row[4]) : new Date();
  const createdAt = Utilities.formatDate(
    createdAtRaw,
    "Asia/Bangkok",
    "dd/MM/yyyy HH:mm:ss"
  );

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  const html = `
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial; padding: 20px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 6px; font-size: 12px; }
          th { background: #f2f2f2; }
        </style>
      </head>
      <body>
        <h2>Order ${escapeHtml(orderId)}</h2>
        <p>Status: ${escapeHtml(status)}</p>
        <p>Date: ${createdAt}</p>

        <table>
          <tr>
            <th>Product</th>
            <th>Qty</th>
            <th>Price</th>
            <th>Total</th>
          </tr>
          ${items.map(i => `
            <tr>
              <td>${escapeHtml(i.name)}</td>
              <td>${i.qty}</td>
              <td>${i.price}</td>
              <td>${Number(i.qty) * Number(i.price)}</td>
            </tr>
          `).join("")}
        </table>

        <h3>Grand Total: ${total}</h3>
        <p>Generated by: ${escapeHtml(by)}</p>
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, "text/html")
    .getAs("application/pdf")
    .setName(`Order-${String(orderId).replace(/[^a-zA-Z0-9-_]/g, "")}.pdf`);

  const file = DriveApp.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE,
    DriveApp.Permission.VIEW
  );

  return {
    url: file.getDownloadUrl(),
    fileId: file.getId()
  };
}

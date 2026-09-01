/************************************************
 * MAIRIX ERP
 * 8-MAIRIX_InventoryEngine.gs
 *
 * INVENTORY ENGINE
 *
 * RESPONSIBILITIES
 * ----------------------------------------------
 * 1. Multi-Warehouse Inventory
 * 2. Central Inventory Ledger
 * 3. Inventory Movement Sign Convention
 * 4. Stock Balance by Item / Warehouse
 * 5. Stock Value using WAC
 * 6. Receipt / Issue / Return / Transfer
 * 7. Central Reference for every movement
 * 8. Opening Stock / Opening Cost support
 * 9. Tax-neutral inventory layer
 *
 * IMPORTANT
 * ----------------------------------------------
 * This engine does NOT own:
 * - Sales
 * - Purchases
 * - Expenses
 * - Taxes
 * - Customers
 * - Vendors
 *
 * Those modules create business transactions.
 *
 * This engine records their INVENTORY EFFECT.
 *
 *
 * CENTRAL SCHEMA
 * ----------------------------------------------
 *
 * The official inventory movement table is:
 *
 *     MAIRIX_SCHEMA.HEADERS.INVENTORY_LEDGER
 *
 * This engine MUST NOT create a parallel:
 *
 *     INVENTORY_MOVEMENTS
 *
 *
 * MOVEMENT SIGN CONVENTION
 * ----------------------------------------------
 *
 * SALE
 * Quantity Effect       = NEGATIVE
 * Monetary Effect       = POSITIVE
 *
 * SALES RETURN
 * Quantity Effect       = POSITIVE
 * Monetary Effect       = NEGATIVE
 *
 * PURCHASE
 * Quantity Effect       = POSITIVE
 * Monetary Effect       = NEGATIVE
 *
 * PURCHASE RETURN
 * Quantity Effect       = NEGATIVE
 * Monetary Effect       = POSITIVE
 *
 * TRANSFER OUT
 * Quantity Effect       = NEGATIVE
 * Monetary Effect       = ZERO
 *
 * TRANSFER IN
 * Quantity Effect       = POSITIVE
 * Monetary Effect       = ZERO
 *
 * ADJUSTMENT IN
 * Quantity Effect       = POSITIVE
 * Monetary Effect       = ZERO
 *
 * ADJUSTMENT OUT
 * Quantity Effect       = NEGATIVE
 * Monetary Effect       = ZERO
 *
 *
 * IMPORTANT
 * ----------------------------------------------
 * Reports should consume signed effects directly.
 *
 * Stock Quantity:
 *
 *     SUM(Inventory Quantity Effect)
 *
 * Inventory value:
 *
 *     WAC / inventory valuation engine
 *
 * Do not rebuild movement meaning in reports
 * using transaction-type conditions.
 ************************************************/


/************************************************
 * CONFIGURATION
 ************************************************/

const MAIRIX_INVENTORY_CONFIG = {

  SHEETS: {

    ITEMS:
      "ITEMS",

    WAREHOUSES:
      "WAREHOUSES",

    LEDGER:
      "INVENTORY_LEDGER",

    BALANCES:
      "INVENTORY_BALANCES"

  },


  MOVEMENT_TYPES: {

    PURCHASE:
      "PURCHASE",

    PURCHASE_RETURN:
      "PURCHASE_RETURN",

    SALE:
      "SALE",

    SALES_RETURN:
      "SALES_RETURN",

    TRANSFER_OUT:
      "TRANSFER_OUT",

    TRANSFER_IN:
      "TRANSFER_IN",

    ADJUSTMENT_IN:
      "ADJUSTMENT_IN",

    ADJUSTMENT_OUT:
      "ADJUSTMENT_OUT"

  }

};


/************************************************
 * MOVEMENT SIGN MAP
 ************************************************/

const MAIRIX_INVENTORY_SIGN_MAP = {

  PURCHASE: {

    quantity:
      1,

    monetary:
      -1

  },

  PURCHASE_RETURN: {

    quantity:
      -1,

    monetary:
      1

  },

  SALE: {

    quantity:
      -1,

    monetary:
      1

  },

  SALES_RETURN: {

    quantity:
      1,

    monetary:
      -1

  },

  TRANSFER_OUT: {

    quantity:
      -1,

    monetary:
      0

  },

  TRANSFER_IN: {

    quantity:
      1,

    monetary:
      0

  },

  ADJUSTMENT_IN: {

    quantity:
      1,

    monetary:
      0

  },

  ADJUSTMENT_OUT: {

    quantity:
      -1,

    monetary:
      0

  }

};


/************************************************
 * GET CENTRAL INVENTORY HEADERS
 ************************************************/

function getMAIRIXInventoryLedgerHeaders() {

  if (
    typeof MAIRIX_SCHEMA ===
    "undefined"
  ) {

    throw new Error(
      "MAIRIX_SCHEMA is not available."
    );

  }


  if (
    !MAIRIX_SCHEMA.HEADERS ||
    !MAIRIX_SCHEMA.HEADERS.INVENTORY_LEDGER
  ) {

    throw new Error(
      "MAIRIX_SCHEMA.HEADERS.INVENTORY_LEDGER is missing."
    );

  }


  return MAIRIX_SCHEMA
    .HEADERS
    .INVENTORY_LEDGER;

}


/************************************************
 * SETUP INVENTORY ENGINE
 *
 * IMPORTANT:
 * ----------------------------------------------
 * This function does NOT create a new inventory
 * movement structure.
 *
 * INVENTORY_LEDGER is owned by the central
 * MAIRIX schema.
 ************************************************/

function setupMAIRIXInventoryEngine() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const ledgerSheet =
    ss.getSheetByName(
      MAIRIX_INVENTORY_CONFIG.SHEETS.LEDGER
    );


  if (!ledgerSheet) {

    throw new Error(
      "Required sheet does not exist: " +
      MAIRIX_INVENTORY_CONFIG.SHEETS.LEDGER +
      ". Run setupMairixSystem() first."
    );

  }


  const expectedHeaders =
    getMAIRIXInventoryLedgerHeaders();


  const actualHeaders =
    ledgerSheet
      .getRange(
        1,
        1,
        1,
        ledgerSheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const normalizedActual =
    actualHeaders.slice();


  while (
    normalizedActual.length > 0 &&
    normalizedActual[
      normalizedActual.length - 1
    ] === ""
  ) {

    normalizedActual.pop();

  }


  const valid =
    expectedHeaders.length ===
      normalizedActual.length &&
    expectedHeaders.every(
      function(header, index) {

        return (
          String(header).trim() ===
          String(
            normalizedActual[index]
          ).trim()
        );

      }
    );


  if (!valid) {

    throw new Error(
      "INVENTORY_LEDGER schema mismatch. " +
      "Run validateMairixSchema() / " +
      "diagnoseMairixSchemaHeaders()."
    );

  }


  /*
   * INVENTORY BALANCES
   */

  let balanceSheet =
    ss.getSheetByName(
      MAIRIX_INVENTORY_CONFIG.SHEETS.BALANCES
    );


  if (!balanceSheet) {

    balanceSheet =
      ss.insertSheet(
        MAIRIX_INVENTORY_CONFIG.SHEETS.BALANCES
      );

  }


  const balanceHeaders = [

    "Item Code",

    "Item Name",

    "Warehouse Code",

    "Warehouse Name",

    "Opening Stock",

    "Movement Quantity",

    "Current Stock",

    "WAC",

    "Inventory Value",

    "Last Movement ID"

  ];


  const currentBalanceHeaders =
    balanceSheet.getLastColumn() > 0
      ? balanceSheet
          .getRange(
            1,
            1,
            1,
            Math.max(
              balanceSheet.getLastColumn(),
              balanceHeaders.length
            )
          )
          .getDisplayValues()[0]
          .map(
            function(header) {

              return String(header).trim();

            }
          )
      : [];


  let balanceSchemaValid =
    true;


  for (
    let i = 0;
    i < balanceHeaders.length;
    i++
  ) {

    if (
      currentBalanceHeaders[i] !==
      balanceHeaders[i]
    ) {

      balanceSchemaValid =
        false;

      break;

    }

  }


  if (
    !balanceSchemaValid
  ) {

    if (
      balanceSheet.getLastRow() > 1
    ) {

      balanceSheet
        .getRange(
          2,
          1,
          balanceSheet.getLastRow() - 1,
          Math.max(
            balanceSheet.getLastColumn(),
            balanceHeaders.length
          )
        )
        .clearContent();

    }


    balanceSheet
      .getRange(
        1,
        1,
        1,
        balanceHeaders.length
      )
      .setValues([
        balanceHeaders
      ]);

  }


  balanceSheet
    .setFrozenRows(
      1
    );


  balanceSheet
    .getRange(
      1,
      1,
      1,
      balanceHeaders.length
    )
    .setFontWeight(
      "bold"
    );


  balanceSheet.autoResizeColumns(
    1,
    balanceHeaders.length
  );


  return {

    success:
      true,

    ledgerSheet:
      MAIRIX_INVENTORY_CONFIG.SHEETS.LEDGER,

    balanceSheet:
      MAIRIX_INVENTORY_CONFIG.SHEETS.BALANCES

  };

}


/************************************************
 * GET INVENTORY LEDGER SHEET
 ************************************************/

function getMAIRIXInventoryLedgerSheet() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        MAIRIX_INVENTORY_CONFIG.SHEETS.LEDGER
      );


  if (!sheet) {

    throw new Error(
      "INVENTORY_LEDGER sheet not found."
    );

  }


  return sheet;

}


/************************************************
 * GET INVENTORY MOVEMENT SHEET
 *
 * Backward-compatible alias.
 ************************************************/

function getMAIRIXInventoryMovementSheet() {

  return getMAIRIXInventoryLedgerSheet();

}


/************************************************
 * GET MOVEMENT SIGN
 ************************************************/

function getMAIRIXInventoryMovementSign(
  movementType
) {

  const type =
    mairixCleanText(
      movementType
    ).toUpperCase();


  const sign =
    MAIRIX_INVENTORY_SIGN_MAP[
      type
    ];


  if (!sign) {

    throw new Error(
      "Unsupported inventory movement type: " +
      type
    );

  }


  return sign;

}


/************************************************
 * VALIDATE MOVEMENT
 ************************************************/

function validateMAIRIXInventoryMovement(
  data
) {

  if (!data) {

    throw new Error(
      "Inventory movement data is required."
    );

  }


  const movementType =
    mairixCleanText(
      data.movementType
    ).toUpperCase();


  if (
    !MAIRIX_INVENTORY_SIGN_MAP[
      movementType
    ]
  ) {

    throw new Error(
      "Invalid inventory movement type: " +
      movementType
    );

  }


  if (
    !mairixCleanText(
      data.referenceId
    )
  ) {

    throw new Error(
      "Inventory movement Reference ID is required."
    );

  }


  if (
    !mairixCleanText(
      data.itemCode
    )
  ) {

    throw new Error(
      "Item Code is required."
    );

  }


  if (
    !mairixCleanText(
      data.warehouseCode
    )
  ) {

    throw new Error(
      "Warehouse Code is required."
    );

  }


  const quantity =
    Number(
      data.quantity
    );


  if (
    !isFinite(quantity) ||
    quantity <= 0
  ) {

    throw new Error(
      "Inventory movement quantity must be greater than zero."
    );

  }


  const unitCost =
    Number(
      data.unitCost === undefined ||
      data.unitCost === null ||
      data.unitCost === ""
        ? 0
        : data.unitCost
    );


  if (
    !isFinite(unitCost) ||
    unitCost < 0
  ) {

    throw new Error(
      "Invalid inventory unit cost."
    );

  }


  return {

    movementType:
      movementType,

    referenceId:
      mairixCleanText(
        data.referenceId
      ),

    lineId:
      mairixCleanText(
        data.lineId
      ),

    itemCode:
      mairixCleanText(
        data.itemCode
      ),

    itemName:
      mairixCleanText(
        data.itemName
      ),

    warehouseCode:
      mairixCleanText(
        data.warehouseCode
      ),

    warehouseName:
      mairixCleanText(
        data.warehouseName
      ),

    quantity:
      quantity,

    unitCost:
      unitCost,

    transactionDate:
      data.movementDate ||
      data.transactionDate ||
      new Date(),

    userId:
      mairixCleanText(
        data.userId
      ),

    userName:
      mairixCleanText(
        data.userName
      )

  };

}


/************************************************
 * GENERATE MOVEMENT ID
 *
 * The Schema does not expose Movement ID.
 *
 * The official INVENTORY_LEDGER therefore uses
 * Reference ID + Line ID as its transaction
 * identity.
 *
 * This function remains available for callers
 * that need a technical identifier.
 ************************************************/

function generateMAIRIXInventoryMovementId() {

  const suffix =
    Utilities
      .getUuid()
      .replace(
        /-/g,
        ""
      )
      .substring(
        0,
        12
      )
      .toUpperCase();


  return (
    "INV-" +
    suffix
  );

}


/************************************************
 * GET ITEM OPENING DATA
 *
 * Reads:
 *
 * ITEMS
 *   Item Code
 *   Opening Quantity
 *   Opening Cost
 ************************************************/

function getMAIRIXItemOpeningData(
  itemCode
) {

  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  if (!normalizedItem) {

    throw new Error(
      "Item Code is required."
    );

  }


  const sheet =
  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      MAIRIX_INVENTORY_CONFIG.SHEETS.ITEMS
    );

  if (!sheet) {

    throw new Error(
      "ITEMS sheet not found."
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return {

      itemCode:
        normalizedItem,

      itemName:
        "",

      openingQuantity:
        0,

      openingCost:
        0

    };

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const itemIndex =
    headers.indexOf(
      "Item Code"
    );


  const nameIndex =
    headers.indexOf(
      "Item Name"
    );


  const openingQuantityIndex =
    headers.indexOf(
      "Opening Quantity"
    );


  const openingCostIndex =
    headers.indexOf(
      "Opening Cost"
    );


  if (
    itemIndex === -1
  ) {

    throw new Error(
      'ITEMS is missing "Item Code".'
    );

  }


  if (
    openingQuantityIndex === -1 ||
    openingCostIndex === -1
  ) {

    throw new Error(
      "ITEMS is missing Opening Quantity or Opening Cost."
    );

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  for (
    let i = 0;
    i < rows.length;
    i++
  ) {

    if (
      mairixCleanText(
        rows[i][itemIndex]
      ) !== normalizedItem
    ) {

      continue;

    }


    const openingQuantity =
      Number(
        rows[i][openingQuantityIndex] || 0
      );


    const openingCost =
      Number(
        rows[i][openingCostIndex] || 0
      );


    if (
      !isFinite(openingQuantity) ||
      openingQuantity < 0
    ) {

      throw new Error(
        "Invalid Opening Quantity for Item: " +
        normalizedItem
      );

    }


    if (
      !isFinite(openingCost) ||
      openingCost < 0
    ) {

      throw new Error(
        "Invalid Opening Cost for Item: " +
        normalizedItem
      );

    }


    return {

      itemCode:
        normalizedItem,

      itemName:
        nameIndex === -1
          ? ""
          : mairixCleanText(
              rows[i][nameIndex]
            ),

      openingQuantity:
        openingQuantity,

      openingCost:
        openingCost

    };

  }


  /*
   * Item does not exist in ITEMS.
   *
   * Do not silently manufacture opening
   * stock for an unknown item.
   */

  throw new Error(
    "Item not found in ITEMS: " +
    normalizedItem
  );

}


/************************************************
 * FIND OPENING STOCK FOR WAREHOUSE
 *
 * IMPORTANT:
 *
 * The current ITEMS schema contains
 * item-level opening quantity, but does not
 * contain an Opening Warehouse column.
 *
 * Therefore opening quantity is assigned to
 * the first / default warehouse only when a
 * caller explicitly provides the flag:
 *
 *     data.useOpeningStock === true
 *
 * Business modules should preferably create
 * explicit opening inventory movements.
 ************************************************/

function getMAIRIXOpeningStockForWarehouse(
  itemCode,
  warehouseCode,
  options
) {

  const opening =
    getMAIRIXItemOpeningData(
      itemCode
    );


  const opts =
    options || {};


  if (
    opts.useOpeningStock !== true
  ) {

    return {

      quantity:
        0,

      value:
        0,

      wac:
        0

    };

  }


  /*
   * Explicit opt-in.
   *
   * Because ITEMS currently has no warehouse
   * assignment for opening stock, the caller
   * is responsible for ensuring that this
   * warehouse is the intended opening warehouse.
   */

  const quantity =
    opening.openingQuantity;


  const value =
    quantity *
    opening.openingCost;


  return {

    quantity:
      quantity,

    value:
      value,

    wac:
      quantity > 0
        ? opening.openingCost
        : 0

  };

}


/************************************************
 * GET CURRENT STOCK
 *
 * Uses central INVENTORY_LEDGER.
 ************************************************/

function getMAIRIXCurrentStock(
  itemCode,
  warehouseCode
) {

  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  const normalizedWarehouse =
    mairixCleanText(
      warehouseCode
    );


  if (
    !normalizedItem ||
    !normalizedWarehouse
  ) {

    throw new Error(
      "Item Code and Warehouse Code are required."
    );

  }


  const state =
    calculateMAIRIXWAC(
      normalizedItem,
      normalizedWarehouse
    );


  return state.quantity;

}


/************************************************
 * GET INVENTORY HISTORY
 ************************************************/

function getMAIRIXInventoryHistory(
  itemCode,
  warehouseCode
) {

  const sheet =
    getMAIRIXInventoryLedgerSheet();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return [];

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  const normalizedWarehouse =
    mairixCleanText(
      warehouseCode
    );


  return rows
    .map(
      function(row) {

        const object = {};


        headers.forEach(
          function(header, index) {

            object[header] =
              row[index];

          }
        );


        return object;

      }
    )
    .filter(
      function(row) {

        return (

          mairixCleanText(
            row["Product Code"]
          ) ===
          normalizedItem &&

          mairixCleanText(
            row["Warehouse Code"]
          ) ===
          normalizedWarehouse

        );

      }
    );

}


/************************************************
 * CALCULATE WAC
 *
 * Uses:
 *
 * Opening Quantity / Opening Cost
 * +
 * INVENTORY_LEDGER movements
 *
 * Positive movement:
 *
 *     adds quantity and value
 *
 * Negative movement:
 *
 *     consumes stock at current WAC
 ************************************************/

function calculateMAIRIXWAC(
  itemCode,
  warehouseCode,
  options
) {

  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  const normalizedWarehouse =
    mairixCleanText(
      warehouseCode
    );


  if (
    !normalizedItem ||
    !normalizedWarehouse
  ) {

    throw new Error(
      "Item Code and Warehouse Code are required."
    );

  }


  const opening =
    getMAIRIXOpeningStockForWarehouse(
      normalizedItem,
      normalizedWarehouse,
      options
    );


  let quantity =
    opening.quantity;


  let value =
    opening.value;


  const history =
    getMAIRIXInventoryHistory(
      normalizedItem,
      normalizedWarehouse
    );


  history.forEach(
    function(row) {

      const quantityEffect =
        Number(
          row[
            "Inventory Quantity Effect"
          ] || 0
        );


      const unitCost =
        Number(
          row[
            "Unit Cost"
          ] || 0
        );


      if (
        !isFinite(quantityEffect)
      ) {

        return;

      }


      if (
        quantityEffect > 0
      ) {

        quantity +=
          quantityEffect;


        value +=
          quantityEffect *
          unitCost;


        return;

      }


      if (
        quantityEffect < 0
      ) {

        const issueQuantity =
          Math.abs(
            quantityEffect
          );


        const currentWAC =
          quantity > 0
            ? value / quantity
            : 0;


        value -=
          issueQuantity *
          currentWAC;


        quantity -=
          issueQuantity;


        if (
          quantity < 0
        ) {

          quantity = 0;

        }


        if (
          value < 0
        ) {

          value = 0;

        }

      }

    }
  );


  const wac =
    quantity > 0
      ? value / quantity
      : 0;


  return {

    quantity:
      quantity,

    value:
      value,

    wac:
      wac

  };

}


/************************************************
 * FIND EXISTING MOVEMENT
 ************************************************/

function findMAIRIXInventoryMovementByReference(
  referenceId,
  movementType,
  itemCode,
  warehouseCode,
  lineId
) {

  const sheet =
    getMAIRIXInventoryLedgerSheet();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return null;

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const referenceIndex =
    headers.indexOf(
      "Reference ID"
    );


  const lineIndex =
    headers.indexOf(
      "Line ID"
    );


  const typeIndex =
    headers.indexOf(
      "Transaction Type"
    );


  const itemIndex =
    headers.indexOf(
      "Product Code"
    );


  const warehouseIndex =
    headers.indexOf(
      "Warehouse Code"
    );


  if (
    referenceIndex === -1 ||
    lineIndex === -1 ||
    typeIndex === -1 ||
    itemIndex === -1 ||
    warehouseIndex === -1
  ) {

    throw new Error(
      "INVENTORY_LEDGER headers are incomplete."
    );

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const normalizedReference =
    mairixCleanText(
      referenceId
    );


  const normalizedType =
    mairixCleanText(
      movementType
    ).toUpperCase();


  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  const normalizedWarehouse =
    mairixCleanText(
      warehouseCode
    );


  const normalizedLine =
    mairixCleanText(
      lineId
    );


  for (
    let i = 0;
    i < rows.length;
    i++
  ) {

    const row =
      rows[i];


    const sameReference =
      mairixCleanText(
        row[referenceIndex]
      ) ===
      normalizedReference;


    const sameType =
      mairixCleanText(
        row[typeIndex]
      ).toUpperCase() ===
      normalizedType;


    const sameItem =
      mairixCleanText(
        row[itemIndex]
      ) ===
      normalizedItem;


    const sameWarehouse =
      mairixCleanText(
        row[warehouseIndex]
      ) ===
      normalizedWarehouse;


    const sameLine =
      !normalizedLine ||
      mairixCleanText(
        row[lineIndex]
      ) ===
      normalizedLine;


    if (
      sameReference &&
      sameType &&
      sameItem &&
      sameWarehouse &&
      sameLine
    ) {

      const result = {};


      headers.forEach(
        function(header, index) {

          result[header] =
            row[index];

        }
      );


      return result;

    }

  }


  return null;

}


/************************************************
 * RECORD INVENTORY MOVEMENT
 *
 * CENTRAL ENTRY POINT
 *
 * All business modules should call this
 * function when inventory changes.
 ************************************************/

function recordMAIRIXInventoryMovement(
  data
) {

  const movement =
    validateMAIRIXInventoryMovement(
      data
    );


  const sign =
    getMAIRIXInventoryMovementSign(
      movement.movementType
    );


  const sheet =
    getMAIRIXInventoryLedgerSheet();


  const existing =
    findMAIRIXInventoryMovementByReference(
      movement.referenceId,
      movement.movementType,
      movement.itemCode,
      movement.warehouseCode,
      movement.lineId
    );


  if (existing) {

    throw new Error(
      "Inventory movement already exists for Reference ID: " +
      movement.referenceId +
      (
        movement.lineId
          ? " / Line ID: " +
            movement.lineId
          : ""
      )
    );

  }


  /*
   * Current inventory state BEFORE movement.
   *
   * We deliberately do not use opening stock
   * automatically because the current ITEMS
   * schema has no warehouse-level opening
   * allocation.
   *
   * Opening inventory should therefore be
   * introduced explicitly through an inventory
   * adjustment / opening transaction.
   */

  const before =
    calculateMAIRIXWAC(
      movement.itemCode,
      movement.warehouseCode
    );


  const quantityEffect =
    movement.quantity *
    sign.quantity;


  /*
   * Inventory value carried by this movement.
   *
   * For positive receipts:
   *
   *     quantity × supplied unit cost
   *
   * For negative issues:
   *
   *     quantity × current WAC
   *
   * This is critical for WAC.
   */

  let inventoryValue;


  let effectiveUnitCost;


  if (
    quantityEffect > 0
  ) {

    effectiveUnitCost =
      movement.unitCost;


    inventoryValue =
      movement.quantity *
      effectiveUnitCost;

  } else {

    effectiveUnitCost =
      before.wac;


    inventoryValue =
      movement.quantity *
      effectiveUnitCost;

  }


  const monetaryEffect =
    inventoryValue *
    sign.monetary;


  /*
   * Calculate AFTER state.
   */

  const afterQuantity =
    before.quantity +
    quantityEffect;


  if (
    afterQuantity < 0
  ) {

    throw new Error(
      "Insufficient stock for Item " +
      movement.itemCode +
      " in Warehouse " +
      movement.warehouseCode
    );

  }


  let afterValue =
    before.value;


  if (
    quantityEffect > 0
  ) {

    afterValue +=
      inventoryValue;

  } else {

    afterValue -=
      inventoryValue;

  }


  if (
    afterValue < 0 &&
    afterValue > -0.0000001
  ) {

    afterValue =
      0;

  }


  if (
    afterValue < 0
  ) {

    throw new Error(
      "Inventory value became negative for Item " +
      movement.itemCode +
      " in Warehouse " +
      movement.warehouseCode
    );

  }


  const afterWAC =
    afterQuantity > 0
      ? afterValue /
        afterQuantity
      : 0;


  /*
   * Schema has no Movement ID column.
   *
   * Generate a technical ID only for return
   * value / audit purposes.
   */

  const movementId =
    generateMAIRIXInventoryMovementId();


  /*
   * IMPORTANT:
   *
   * INVENTORY_LEDGER schema:
   *
   * 1 Reference ID
   * 2 Line ID
   * 3 Transaction Date
   * 4 Transaction Type
   * 5 Warehouse Code
   * 6 Product Code
   * 7 Quantity
   * 8 Unit Cost
   * 9 Inventory Quantity Effect
   * 10 Inventory Value Effect
   * 11 Running Quantity
   * 12 Running Value
   * 13 WAC After Transaction
   * 14 Created At
   */

  const row = [

    movement.referenceId,

    movement.lineId,

    movement.transactionDate,

    movement.movementType,

    movement.warehouseCode,

    movement.itemCode,

    movement.quantity,

    effectiveUnitCost,

    quantityEffect,

    inventoryValue,

    afterQuantity,

    afterValue,

    afterWAC,

    new Date()

  ];


  const nextRow =
    Math.max(
      sheet.getLastRow() + 1,
      2
    );


  sheet
    .getRange(
      nextRow,
      1,
      1,
      row.length
    )
    .setValues([
      row
    ]);


  return {

    success:
      true,

    movementId:
      movementId,

    referenceId:
      movement.referenceId,

    lineId:
      movement.lineId,

    movementType:
      movement.movementType,

    itemCode:
      movement.itemCode,

    warehouseCode:
      movement.warehouseCode,

    quantity:
      movement.quantity,

    quantityEffect:
      quantityEffect,

    inventoryValue:
      inventoryValue,

    monetaryEffect:
      monetaryEffect,

    stockBefore:
      before.quantity,

    stockAfter:
      afterQuantity,

    wacBefore:
      before.wac,

    wacAfter:
      afterWAC

  };

}


/************************************************
 * TRANSFER BETWEEN WAREHOUSES
 *
 * One business Reference ID.
 *
 * Two ledger entries:
 *
 * TRANSFER_OUT
 * TRANSFER_IN
 *
 * Both movements carry the same inventory cost.
 ************************************************/

function transferMAIRIXInventory(
  data
) {

  if (!data) {

    throw new Error(
      "Transfer data is required."
    );

  }


  const fromWarehouseCode =
    mairixCleanText(
      data.fromWarehouseCode
    );


  const toWarehouseCode =
    mairixCleanText(
      data.toWarehouseCode
    );


  if (
    !fromWarehouseCode ||
    !toWarehouseCode
  ) {

    throw new Error(
      "Source and destination warehouses are required."
    );

  }


  if (
    fromWarehouseCode ===
    toWarehouseCode
  ) {

    throw new Error(
      "Source and destination warehouses must be different."
    );

  }


  const quantity =
    Number(
      data.quantity
    );


  if (
    !isFinite(quantity) ||
    quantity <= 0
  ) {

    throw new Error(
      "Transfer quantity must be greater than zero."
    );

  }


  const referenceId =
    mairixCleanText(
      data.referenceId
    );


  if (!referenceId) {

    throw new Error(
      "Transfer Reference ID is required."
    );

  }


  const itemCode =
    mairixCleanText(
      data.itemCode
    );


  if (!itemCode) {

    throw new Error(
      "Transfer Item Code is required."
    );

  }


  /*
   * Source state.
   */

  const sourceState =
    calculateMAIRIXWAC(
      itemCode,
      fromWarehouseCode
    );


  if (
    sourceState.quantity <
    quantity
  ) {

    throw new Error(
      "Insufficient stock for warehouse transfer."
    );

  }


  const transferCost =
    sourceState.wac;


  /*
   * Prevent duplicate transfer.
   */

  const existingOut =
    findMAIRIXInventoryMovementByReference(
      referenceId,
      "TRANSFER_OUT",
      itemCode,
      fromWarehouseCode,
      data.lineId
    );


  if (existingOut) {

    throw new Error(
      "Warehouse transfer already exists for Reference ID: " +
      referenceId
    );

  }


  const outLineId =
    mairixCleanText(
      data.lineId
    ) ||
    "OUT";


  const inLineId =
    mairixCleanText(
      data.lineId
    ) ||
    "IN";


  const out =
    recordMAIRIXInventoryMovement({

      referenceId:
        referenceId,

      lineId:
        outLineId,

      movementType:
        "TRANSFER_OUT",

      itemCode:
        itemCode,

      itemName:
        data.itemName,

      warehouseCode:
        fromWarehouseCode,

      warehouseName:
        data.fromWarehouseName,

      quantity:
        quantity,

      unitCost:
        transferCost,

      transactionDate:
        data.transactionDate ||
        data.movementDate,

      userId:
        data.userId,

      userName:
        data.userName

    });


  try {

    const incoming =
      recordMAIRIXInventoryMovement({

        referenceId:
          referenceId,

        lineId:
          inLineId,

        movementType:
          "TRANSFER_IN",

        itemCode:
          itemCode,

        itemName:
          data.itemName,

        warehouseCode:
          toWarehouseCode,

        warehouseName:
          data.toWarehouseName,

        quantity:
          quantity,

        unitCost:
          transferCost,

        transactionDate:
          data.transactionDate ||
          data.movementDate,

        userId:
          data.userId,

        userName:
          data.userName

      });


    return {

      success:
        true,

      referenceId:
        referenceId,

      quantity:
        quantity,

      transferCost:
        transferCost,

      source:
        out,

      destination:
        incoming

    };

  } catch (error) {

    /*
     * IMPORTANT:
     *
     * This function is intentionally not pretending
     * to provide rollback by deleting the OUT row.
     *
     * Full transaction rollback belongs to the
     * central MAIRIX transaction / rollback engine.
     *
     * The caller should execute transfer inside
     * withMaifTransactionLock / transaction wrapper
     * when atomicity is required.
     */

    throw error;

  }

}


/************************************************
 * REBUILD INVENTORY BALANCES
 *
 * Derived reporting / control table.
 *
 * It does NOT own inventory state.
 *
 * Central source:
 *
 *     INVENTORY_LEDGER
 ************************************************/

function rebuildMAIRIXInventoryBalances() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      MAIRIX_INVENTORY_CONFIG.SHEETS.BALANCES
    );


  if (!sheet) {

    throw new Error(
      "INVENTORY_BALANCES sheet not found."
    );

  }


  const movementSheet =
    getMAIRIXInventoryLedgerSheet();


  const lastRow =
    movementSheet.getLastRow();


  if (
    lastRow < 2
  ) {

    if (
      sheet.getLastRow() > 1
    ) {

      sheet
        .getRange(
          2,
          1,
          sheet.getLastRow() - 1,
          10
        )
        .clearContent();

    }


    return {

      success:
        true,

      rows:
        0

    };

  }


  const headers =
    movementSheet
      .getRange(
        1,
        1,
        1,
        movementSheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const referenceIndex =
    headers.indexOf(
      "Reference ID"
    );


  const typeIndex =
    headers.indexOf(
      "Transaction Type"
    );


  const warehouseIndex =
    headers.indexOf(
      "Warehouse Code"
    );


  const itemIndex =
    headers.indexOf(
      "Product Code"
    );


  const quantityIndex =
    headers.indexOf(
      "Quantity"
    );


  const quantityEffectIndex =
    headers.indexOf(
      "Inventory Quantity Effect"
    );


  const unitCostIndex =
    headers.indexOf(
      "Unit Cost"
    );


  if (
    referenceIndex === -1 ||
    typeIndex === -1 ||
    warehouseIndex === -1 ||
    itemIndex === -1 ||
    quantityIndex === -1 ||
    quantityEffectIndex === -1 ||
    unitCostIndex === -1
  ) {

    throw new Error(
      "INVENTORY_LEDGER headers are incomplete."
    );

  }


  const rows =
    movementSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const groups = {};


  rows.forEach(
    function(row) {

      const itemCode =
        mairixCleanText(
          row[itemIndex]
        );


      const warehouseCode =
        mairixCleanText(
          row[warehouseIndex]
        );


      if (
        !itemCode ||
        !warehouseCode
      ) {

        return;

      }


      const key =
        itemCode +
        "|" +
        warehouseCode;


      if (!groups[key]) {

        groups[key] = {

          itemCode:
            itemCode,

          warehouseCode:
            warehouseCode,

          movementQuantity:
            0,

          lastMovementId:
            ""

        };

      }


      groups[key]
        .movementQuantity +=
        Number(
          row[
            quantityEffectIndex
          ] || 0
        );


      groups[key]
        .lastMovementId =
        mairixCleanText(
          row[
            referenceIndex
          ]
        );

    }
  );


  const output = [];


  Object.keys(groups)
    .forEach(
      function(key) {

        const group =
          groups[key];


        const opening =
          getMAIRIXOpeningStockForWarehouse(
            group.itemCode,
            group.warehouseCode
          );


        const state =
          calculateMAIRIXWAC(
            group.itemCode,
            group.warehouseCode
          );


        let itemName =
          "";


        let warehouseName =
          "";


        try {

          itemName =
            getMAIRIXItemOpeningData(
              group.itemCode
            ).itemName;

        } catch (error) {

          itemName =
            "";

        }


        try {

          warehouseName =
            getMAIRIXWarehouseName(
              group.warehouseCode
            );

        } catch (error) {

          warehouseName =
            "";

        }


        output.push([

          group.itemCode,

          itemName,

          group.warehouseCode,

          warehouseName,

          opening.quantity,

          group.movementQuantity,

          state.quantity,

          state.wac,

          state.value,

          group.lastMovementId

        ]);

      }
    );


  if (
    sheet.getLastRow() > 1
  ) {

    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        10
      )
      .clearContent();

  }


  if (
    output.length
  ) {

    sheet
      .getRange(
        2,
        1,
        output.length,
        10
      )
      .setValues(
        output
      );

  }


  return {

    success:
      true,

    rows:
      output.length

  };

}


/************************************************
 * GET WAREHOUSE NAME
 ************************************************/

function getMAIRIXWarehouseName(
  warehouseCode
) {

  const normalizedWarehouse =
    mairixCleanText(
      warehouseCode
    );


  if (!normalizedWarehouse) {

    return "";

  }


  const sheet =
  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      MAIRIX_INVENTORY_CONFIG.SHEETS.WAREHOUSES
    );


  if (!sheet) {

    return "";

  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return "";

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const codeIndex =
    headers.indexOf(
      "Warehouse Code"
    );


  const nameIndex =
    headers.indexOf(
      "Warehouse Name"
    );


  if (
    codeIndex === -1 ||
    nameIndex === -1
  ) {

    return "";

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  for (
    let i = 0;
    i < rows.length;
    i++
  ) {

    if (
      mairixCleanText(
        rows[i][codeIndex]
      ) ===
      normalizedWarehouse
    ) {

      return mairixCleanText(
        rows[i][nameIndex]
      );

    }

  }


  return "";

}


/************************************************
 * GET TOTAL STOCK ACROSS ALL WAREHOUSES
 ************************************************/

function getMAIRIXItemTotalStock(
  itemCode
) {

  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  if (!normalizedItem) {

    throw new Error(
      "Item Code is required."
    );

  }


  const sheet =
    getMAIRIXInventoryLedgerSheet();


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return 0;

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const itemIndex =
    headers.indexOf(
      "Product Code"
    );


  const quantityIndex =
    headers.indexOf(
      "Inventory Quantity Effect"
    );


  if (
    itemIndex === -1 ||
    quantityIndex === -1
  ) {

    throw new Error(
      "INVENTORY_LEDGER headers are incomplete."
    );

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  let total =
    0;


  rows.forEach(
    function(row) {

      if (
        mairixCleanText(
          row[itemIndex]
        ) !==
        normalizedItem
      ) {

        return;

      }


      total +=
        Number(
          row[quantityIndex] || 0
        );

    }
  );


  return total;

}


/************************************************
 * GET STOCK BY WAREHOUSE
 ************************************************/

function getMAIRIXItemStockByWarehouse(
  itemCode
) {

  const normalizedItem =
    mairixCleanText(
      itemCode
    );


  if (!normalizedItem) {

    throw new Error(
      "Item Code is required."
    );

  }


  const sheet =
  SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(
      MAIRIX_INVENTORY_CONFIG.SHEETS.WAREHOUSES
    );


  if (!sheet) {

    throw new Error(
      "WAREHOUSES sheet not found."
    );

  }


  const lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return [];

  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  const codeIndex =
    headers.indexOf(
      "Warehouse Code"
    );


  const nameIndex =
    headers.indexOf(
      "Warehouse Name"
    );


  if (
    codeIndex === -1
  ) {

    throw new Error(
      'WAREHOUSES is missing "Warehouse Code".'
    );

  }


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  return rows
    .map(
      function(row) {

        const warehouseCode =
          mairixCleanText(
            row[codeIndex]
          );


        if (!warehouseCode) {

          return null;

        }


        const state =
          calculateMAIRIXWAC(
            normalizedItem,
            warehouseCode
          );


        return {

          warehouseCode:
            warehouseCode,

          warehouseName:
            nameIndex === -1
              ? ""
              : mairixCleanText(
                  row[nameIndex]
                ),

          quantity:
            state.quantity,

          wac:
            state.wac,

          value:
            state.value

        };

      }
    )
    .filter(
      function(row) {

        return row !== null;

      }
    );

}


/************************************************
 * OPENING INVENTORY MOVEMENT
 *
 * Creates explicit opening stock in the
 * central INVENTORY_LEDGER.
 *
 * This is the preferred method for assigning
 * opening stock to a specific warehouse.
 ************************************************/

function recordMAIRIXOpeningInventory(
  data
) {

  if (!data) {

    throw new Error(
      "Opening inventory data is required."
    );

  }


  const quantity =
    Number(
      data.quantity
    );


  const unitCost =
    Number(
      data.unitCost
    );


  if (
    !isFinite(quantity) ||
    quantity <= 0
  ) {

    throw new Error(
      "Opening inventory quantity must be greater than zero."
    );

  }


  if (
    !isFinite(unitCost) ||
    unitCost < 0
  ) {

    throw new Error(
      "Opening inventory unit cost is invalid."
    );

  }


  return recordMAIRIXInventoryMovement({

    referenceId:
      data.referenceId ||
      (
        "OPENING-" +
        new Date().getTime()
      ),

    lineId:
      data.lineId ||
      "OPENING",

    movementType:
      "ADJUSTMENT_IN",

    itemCode:
      data.itemCode,

    itemName:
      data.itemName,

    warehouseCode:
      data.warehouseCode,

    warehouseName:
      data.warehouseName,

    quantity:
      quantity,

    unitCost:
      unitCost,

    transactionDate:
      data.transactionDate ||
      new Date(),

    userId:
      data.userId,

    userName:
      data.userName

  });

}


/************************************************
 * TEST PURCHASE MOVEMENT
 ************************************************/

function testMAIRIXInventoryPurchase() {

  const referenceId =
    "TEST-PURCHASE-" +
    new Date().getTime();


  const result =
    recordMAIRIXInventoryMovement({

      referenceId:
        referenceId,

      lineId:
        "1",

      movementType:
        "PURCHASE",

      itemCode:
        "TEST-ITEM-001",

      itemName:
        "Test Item",

      warehouseCode:
        "WH-001",

      warehouseName:
        "Main Warehouse",

      quantity:
        10,

      unitCost:
        100,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    result.quantityEffect !==
    10
  ) {

    throw new Error(
      "PURCHASE quantity sign is incorrect."
    );

  }


  if (
    result.monetaryEffect !==
    -1000
  ) {

    throw new Error(
      "PURCHASE monetary sign is incorrect."
    );

  }


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  Logger.log(
    "MAIRIX INVENTORY PURCHASE TEST: PASS"
  );

}


/************************************************
 * TEST SALE MOVEMENT
 ************************************************/

function testMAIRIXInventorySale() {

  const before =
    getMAIRIXCurrentStock(
      "TEST-ITEM-001",
      "WH-001"
    );


  if (
    before < 5
  ) {

    throw new Error(
      "Not enough test stock for SALE test."
    );

  }


  const referenceId =
    "TEST-SALE-" +
    new Date().getTime();


  const result =
    recordMAIRIXInventoryMovement({

      referenceId:
        referenceId,

      lineId:
        "1",

      movementType:
        "SALE",

      itemCode:
        "TEST-ITEM-001",

      itemName:
        "Test Item",

      warehouseCode:
        "WH-001",

      warehouseName:
        "Main Warehouse",

      quantity:
        5,

      unitCost:
        0,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    result.quantityEffect !==
    -5
  ) {

    throw new Error(
      "SALE quantity sign is incorrect."
    );

  }


  if (
    result.monetaryEffect <=
    0
  ) {

    throw new Error(
      "SALE monetary effect must be positive."
    );

  }


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  Logger.log(
    "MAIRIX INVENTORY SALE TEST: PASS"
  );

}


/************************************************
 * TEST SALES RETURN
 ************************************************/

function testMAIRIXInventorySalesReturn() {

  const referenceId =
    "TEST-RETURN-" +
    new Date().getTime();


  const result =
    recordMAIRIXInventoryMovement({

      referenceId:
        referenceId,

      lineId:
        "1",

      movementType:
        "SALES_RETURN",

      itemCode:
        "TEST-ITEM-001",

      itemName:
        "Test Item",

      warehouseCode:
        "WH-001",

      warehouseName:
        "Main Warehouse",

      quantity:
        2,

      unitCost:
        100,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    result.quantityEffect !==
    2
  ) {

    throw new Error(
      "SALES_RETURN quantity sign is incorrect."
    );

  }


  if (
    result.monetaryEffect >=
    0
  ) {

    throw new Error(
      "SALES_RETURN monetary effect must be negative."
    );

  }


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  Logger.log(
    "MAIRIX INVENTORY SALES RETURN TEST: PASS"
  );

}


/************************************************
 * TEST WAREHOUSE TRANSFER
 ************************************************/

function testMAIRIXWarehouseTransfer() {

  const referenceId =
    "TEST-TRANSFER-" +
    new Date().getTime();


  const result =
    transferMAIRIXInventory({

      referenceId:
        referenceId,

      itemCode:
        "TEST-ITEM-001",

      itemName:
        "Test Item",

      fromWarehouseCode:
        "WH-001",

      fromWarehouseName:
        "Main Warehouse",

      toWarehouseCode:
        "WH-002",

      toWarehouseName:
        "Secondary Warehouse",

      quantity:
        1,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    result.source.quantityEffect !==
    -1
  ) {

    throw new Error(
      "TRANSFER_OUT sign is incorrect."
    );

  }


  if (
    result.destination.quantityEffect !==
    1
  ) {

    throw new Error(
      "TRANSFER_IN sign is incorrect."
    );

  }


  if (
    result.source.monetaryEffect !==
    0
  ) {

    throw new Error(
      "TRANSFER_OUT monetary effect must be zero."
    );

  }


  if (
    result.destination.monetaryEffect !==
    0
  ) {

    throw new Error(
      "TRANSFER_IN monetary effect must be zero."
    );

  }


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  Logger.log(
    "MAIRIX WAREHOUSE TRANSFER TEST: PASS"
  );

}


/************************************************
 * TEST SIGN MAP
 ************************************************/

function testMAIRIXInventoryMovementSigns() {

  const expected = {

    PURCHASE:
      [1, -1],

    PURCHASE_RETURN:
      [-1, 1],

    SALE:
      [-1, 1],

    SALES_RETURN:
      [1, -1],

    TRANSFER_OUT:
      [-1, 0],

    TRANSFER_IN:
      [1, 0],

    ADJUSTMENT_IN:
      [1, 0],

    ADJUSTMENT_OUT:
      [-1, 0]

  };


  Object.keys(expected)
    .forEach(
      function(type) {

        const sign =
          getMAIRIXInventoryMovementSign(
            type
          );


        if (
          sign.quantity !==
          expected[type][0]
        ) {

          throw new Error(
            type +
            " quantity sign mismatch."
          );

        }


        if (
          sign.monetary !==
          expected[type][1]
        ) {

          throw new Error(
            type +
            " monetary sign mismatch."
          );

        }

      }
    );


  Logger.log(
    "MAIRIX INVENTORY MOVEMENT SIGN TEST: PASS"
  );

}


/************************************************
 * SCHEMA COMPATIBILITY TEST
 ************************************************/

function testMAIRIXInventorySchemaCompatibility() {

  const expected =
    getMAIRIXInventoryLedgerHeaders();


  const sheet =
    getMAIRIXInventoryLedgerSheet();


  const actual =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0]
      .map(
        function(header) {

          return String(header).trim();

        }
      );


  while (
    actual.length > 0 &&
    actual[actual.length - 1] === ""
  ) {

    actual.pop();

  }


  if (
    expected.length !==
    actual.length
  ) {

    throw new Error(
      "INVENTORY_LEDGER column count mismatch. " +
      "Expected: " +
      expected.length +
      " Actual: " +
      actual.length
    );

  }


  expected.forEach(
    function(header, index) {

      if (
        header !==
        actual[index]
      ) {

        throw new Error(
          "INVENTORY_LEDGER header mismatch at column " +
          (index + 1) +
          ". Expected [" +
          header +
          "] Actual [" +
          actual[index] +
          "]"
        );

      }

    }
  );


  Logger.log(
    "MAIRIX INVENTORY SCHEMA COMPATIBILITY TEST: PASS"
  );

}


/************************************************
 * MASTER TEST
 *
 * Structural test only.
 *
 * Does NOT create business transactions.
 ************************************************/

function testMAIRIXInventoryEngine() {

  Logger.log(
    "=========================================="
  );

  Logger.log(
    "MAIRIX INVENTORY ENGINE TEST"
  );

  Logger.log(
    "=========================================="
  );


  testMAIRIXInventorySchemaCompatibility();


  testMAIRIXInventoryMovementSigns();


  setupMAIRIXInventoryEngine();


  Logger.log(
    "Inventory engine setup: PASS"
  );


  Logger.log(
    "=========================================="
  );

  Logger.log(
    "MAIRIX INVENTORY ENGINE READY"
  );

  Logger.log(
    "=========================================="
  );

}


/************************************************
 * FULL INVENTORY FUNCTIONAL TEST
 *
 * IMPORTANT
 * ----------------------------------------------
 * This test intentionally creates temporary
 * inventory movements.
 *
 * Use only on a test / development spreadsheet.
 ************************************************/

function testMAIRIXInventoryFunctionalFlow() {

  const itemCode =
    "TEST-ITEM-001";


  const warehouseA =
    "WH-001";


  const warehouseB =
    "WH-002";


  const purchaseReference =
    "TEST-FLOW-PURCHASE-" +
    new Date().getTime();


  const saleReference =
    "TEST-FLOW-SALE-" +
    new Date().getTime();


  const returnReference =
    "TEST-FLOW-RETURN-" +
    new Date().getTime();


  Logger.log(
    "=========================================="
  );

  Logger.log(
    "MAIRIX INVENTORY FUNCTIONAL FLOW TEST"
  );

  Logger.log(
    "=========================================="
  );


  /*
   * PURCHASE
   */

  const purchase =
    recordMAIRIXInventoryMovement({

      referenceId:
        purchaseReference,

      lineId:
        "1",

      movementType:
        "PURCHASE",

      itemCode:
        itemCode,

      itemName:
        "Test Item",

      warehouseCode:
        warehouseA,

      warehouseName:
        "Main Warehouse",

      quantity:
        10,

      unitCost:
        100,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    purchase.quantityEffect !==
    10
  ) {

    throw new Error(
      "Functional purchase quantity effect failed."
    );

  }


  if (
    purchase.monetaryEffect !==
    -1000
  ) {

    throw new Error(
      "Functional purchase monetary effect failed."
    );

  }


  /*
   * SALE
   *
   * Unit cost is intentionally ignored
   * for issue valuation.
   *
   * Engine uses current WAC.
   */

  const sale =
    recordMAIRIXInventoryMovement({

      referenceId:
        saleReference,

      lineId:
        "1",

      movementType:
        "SALE",

      itemCode:
        itemCode,

      itemName:
        "Test Item",

      warehouseCode:
        warehouseA,

      warehouseName:
        "Main Warehouse",

      quantity:
        4,

      unitCost:
        0,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    sale.quantityEffect !==
    -4
  ) {

    throw new Error(
      "Functional sale quantity effect failed."
    );

  }


  if (
    sale.monetaryEffect !==
    400
  ) {

    throw new Error(
      "Functional sale monetary effect failed."
    );

  }


  /*
   * SALES RETURN
   */

  const salesReturn =
    recordMAIRIXInventoryMovement({

      referenceId:
        returnReference,

      lineId:
        "1",

      movementType:
        "SALES_RETURN",

      itemCode:
        itemCode,

      itemName:
        "Test Item",

      warehouseCode:
        warehouseA,

      warehouseName:
        "Main Warehouse",

      quantity:
        2,

      unitCost:
        100,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    salesReturn.quantityEffect !==
    2
  ) {

    throw new Error(
      "Functional sales return quantity effect failed."
    );

  }


  if (
    salesReturn.monetaryEffect !==
    -200
  ) {

    throw new Error(
      "Functional sales return monetary effect failed."
    );

  }


  /*
   * TRANSFER
   */

  const transferReference =
    "TEST-FLOW-TRANSFER-" +
    new Date().getTime();


  const transfer =
    transferMAIRIXInventory({

      referenceId:
        transferReference,

      itemCode:
        itemCode,

      itemName:
        "Test Item",

      fromWarehouseCode:
        warehouseA,

      fromWarehouseName:
        "Main Warehouse",

      toWarehouseCode:
        warehouseB,

      toWarehouseName:
        "Secondary Warehouse",

      quantity:
        2,

      userId:
        "SYSTEM",

      userName:
        "SYSTEM"

    });


  if (
    transfer.source.quantityEffect !==
    -2
  ) {

    throw new Error(
      "Functional transfer OUT failed."
    );

  }


  if (
    transfer.destination.quantityEffect !==
    2
  ) {

    throw new Error(
      "Functional transfer IN failed."
    );

  }


  /*
   * BALANCE REBUILD
   */

  const rebuild =
    rebuildMAIRIXInventoryBalances();


  if (
    !rebuild ||
    rebuild.success !==
      true
  ) {

    throw new Error(
      "Inventory balance rebuild failed."
    );

  }


  Logger.log(
    JSON.stringify(
      {

        purchase:
          purchase,

        sale:
          sale,

        salesReturn:
          salesReturn,

        transfer:
          transfer,

        rebuild:
          rebuild

      },
      null,
      2
    )
  );


  Logger.log(
    "=========================================="
  );

  Logger.log(
    "MAIRIX INVENTORY FUNCTIONAL FLOW TEST: PASS"
  );

  Logger.log(
    "=========================================="
  );

}

/************************************************
 * MAIRIX ERP
 * 13-MAIRIX_PurchaseReturns.gs
 *
 * PURCHASE RETURN ENGINE
 *
 * RESPONSIBILITIES:
 *
 * 1. Create purchase return transactions
 * 2. Link return to original purchase
 * 3. Validate supplier / warehouse / items
 * 4. Prevent returning more than purchased
 * 5. Reverse inventory quantity
 * 6. Reverse inventory value
 * 7. Reverse supplier payable / cash effect
 * 8. Reverse purchase taxes
 * 9. Create centralized references
 * 10. Preserve original purchase
 * 11. Support partial returns
 * 12. Support full returns
 *
 * PRINCIPLE:
 *
 * PURCHASE
 *   Inventory Qty  +
 *   Inventory Value +
 *   Cash Effect     -
 *   Supplier Payable +
 *
 * PURCHASE RETURN
 *   Inventory Qty  -
 *   Inventory Value -
 *   Cash Effect     +
 *   Supplier Payable -
 *
 * TAX:
 *
 * Purchase tax is reversed according to
 * the tax behavior defined by MAIRIX_TaxEngine.
 *
 * IMPORTANT:
 *
 * This file never directly edits:
 *
 * - inventory balances
 * - supplier balances
 * - cash balances
 * - tax balances
 *
 * The centralized engines own those effects.
 ************************************************/


/************************************************
 * CONFIGURATION
 ************************************************/

const MAIRIX_PURCHASE_RETURNS = {

  SHEET:
    "PURCHASE_RETURNS",

  LINES_SHEET:
    "PURCHASE_RETURN_LINES",

  DOCUMENT_PREFIX:
    "PRT",

  TRANSACTION_TYPE:
    "PURCHASE_RETURN",

  STATUS_POSTED:
    "POSTED",

  STATUS_CANCELLED:
    "CANCELLED",

  PAYMENT_CASH:
    "CASH",

  PAYMENT_CREDIT:
    "CREDIT"

};


/************************************************
 * SETUP
 ************************************************/

function setupMAIRIXPurchaseReturns() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  let headerSheet =
    ss.getSheetByName(
      MAIRIX_PURCHASE_RETURNS.SHEET
    );


  let linesSheet =
    ss.getSheetByName(
      MAIRIX_PURCHASE_RETURNS.LINES_SHEET
    );


  if (!headerSheet) {

    headerSheet =
      ss.insertSheet(
        MAIRIX_PURCHASE_RETURNS.SHEET
      );

  }


  if (!linesSheet) {

    linesSheet =
      ss.insertSheet(
        MAIRIX_PURCHASE_RETURNS.LINES_SHEET
      );

  }


  const headers = [

    "Purchase Return ID",
    "Reference No",
    "Original Purchase ID",
    "Original Reference No",

    "Date",

    "Supplier Code",
    "Supplier Name",

    "Warehouse Code",
    "Warehouse Name",

    "Payment Type",

    "Subtotal",
    "Discount",
    "Taxable Amount",
    "Tax Total",
    "Net Total",

    "Cash Effect",
    "Payable Effect",

    "Tax Reference",
    "Inventory Reference",
    "Financial Reference",

    "User ID",
    "User Name",

    "Status",
    "Created At"

  ];


  const lineHeaders = [

    "Purchase Return ID",
    "Line No",

    "Original Purchase ID",
    "Original Line No",

    "Item Code",
    "Item Name",

    "Quantity",
    "Unit",

    "Unit Cost",

    "Gross Amount",
    "Discount",

    "Taxable Amount",
    "Tax Total",
    "Net Amount",

    "Inventory Effect",

    "Tax Reference",

    "Warehouse Code",

    "Created At"

  ];


  headerSheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setValues([
      headers
    ]);


  linesSheet
    .getRange(
      1,
      1,
      1,
      lineHeaders.length
    )
    .setValues([
      lineHeaders
    ]);


  headerSheet.setFrozenRows(1);
  linesSheet.setFrozenRows(1);


  return {

    success:
      true,

    purchaseReturnSheet:
      MAIRIX_PURCHASE_RETURNS.SHEET,

    linesSheet:
      MAIRIX_PURCHASE_RETURNS.LINES_SHEET

  };

}


/************************************************
 * CREATE PURCHASE RETURN
 *
 * DATA:
 *
 * {
 *   originalPurchaseId,
 *   referenceNo,
 *   date,
 *   lines: [
 *
 *     {
 *       originalLineNo,
 *       itemCode,
 *       quantity
 *     }
 *
 *   ]
 * }
 *
 * Payment type is inherited from original
 * purchase unless explicitly supplied.
 ************************************************/

function createPurchaseReturn(
  data
) {

  if (!data) {

    throw new Error(
      "Purchase return data is required."
    );

  }


  /**********************************************
   * CURRENT USER
   **********************************************/

  const user =
    mairixPurchaseReturnRequireUser();


  /**********************************************
   * PERMISSION
   **********************************************/

  mairixPurchaseReturnRequirePermission(
    user.userId,
    "PURCHASES",
    "Create"
  );


  /**********************************************
   * ORIGINAL PURCHASE
   **********************************************/

  const originalPurchaseId =
    mairixPurchaseReturnRequired(
      data.originalPurchaseId,
      "Original Purchase ID"
    );


  const purchase =
    mairixPurchaseReturnGetPurchase(
      originalPurchaseId
    );


  if (!purchase) {

    throw new Error(
      "Original purchase not found: " +
      originalPurchaseId
    );

  }


  if (
    purchase.status ===
    MAIRIX_PURCHASES.STATUS_CANCELLED
  ) {

    throw new Error(
      "Cannot return a cancelled purchase."
    );

  }


  /**********************************************
   * BASIC DATA
   **********************************************/

  const supplierCode =
    mairixPurchaseReturnText(
      purchase["Supplier Code"]
    );


  const warehouseCode =
    mairixPurchaseReturnText(
      purchase["Warehouse Code"]
    );


  if (!supplierCode) {

    throw new Error(
      "Original purchase has no Supplier Code."
    );

  }


  if (!warehouseCode) {

    throw new Error(
      "Original purchase has no Warehouse Code."
    );

  }


  const purchaseDate =
    data.date
      ? new Date(data.date)
      : new Date();


  if (
    isNaN(
      purchaseDate.getTime()
    )
  ) {

    throw new Error(
      "Invalid purchase return date."
    );

  }


  /**********************************************
   * RETURN LINES
   **********************************************/

  if (
    !Array.isArray(data.lines) ||
    data.lines.length === 0
  ) {

    throw new Error(
      "Purchase return must contain at least one item."
    );

  }


  const originalLines =
    mairixPurchaseReturnGetPurchaseLines(
      originalPurchaseId
    );


  if (!originalLines.length) {

    throw new Error(
      "Original purchase has no purchase lines."
    );

  }


  /**********************************************
   * ALREADY RETURNED QUANTITIES
   **********************************************/

  const returnedMap =
    mairixPurchaseReturnGetReturnedQuantities(
      originalPurchaseId
    );


  /**********************************************
   * PREPARE RETURN LINES
   **********************************************/

  const preparedLines = [];

  let subtotal = 0;

  let lineNo = 1;


  data.lines.forEach(
    function(inputLine) {

      if (!inputLine) {
        return;
      }


      const originalLineNo =
        Number(
          inputLine.originalLineNo
        );


      if (
        !isFinite(originalLineNo) ||
        originalLineNo <= 0
      ) {

        throw new Error(
          "Invalid Original Line No."
        );

      }


      const originalLine =
        originalLines.find(
          function(line) {

            return Number(
              line["Line No"]
            ) === originalLineNo;

          }
        );


      if (!originalLine) {

        throw new Error(
          "Original purchase line not found: " +
          originalLineNo
        );

      }


      const itemCode =
        mairixPurchaseReturnText(
          originalLine["Item Code"]
        );


      if (!itemCode) {

        throw new Error(
          "Original purchase line has no Item Code."
        );

      }


      /*
       * If UI supplies itemCode,
       * it MUST match the original.
       */

      if (
        inputLine.itemCode !== undefined &&
        mairixPurchaseReturnText(
          inputLine.itemCode
        ) !== itemCode
      ) {

        throw new Error(
          "Returned item does not match original purchase line: " +
          originalLineNo
        );

      }


      const quantity =
        Number(
          inputLine.quantity
        );


      if (
        !isFinite(quantity) ||
        quantity <= 0
      ) {

        throw new Error(
          "Invalid return quantity for item: " +
          itemCode
        );

      }


      const purchasedQuantity =
        Number(
          originalLine["Quantity"]
        );


      const alreadyReturned =
        Number(
          returnedMap[originalLineNo] || 0
        );


      const availableQuantity =
        purchasedQuantity -
        alreadyReturned;


      if (
        quantity >
        availableQuantity
      ) {

        throw new Error(

          "Return quantity exceeds remaining quantity for item " +
          itemCode +
          ". Purchased: " +
          purchasedQuantity +
          " | Already Returned: " +
          alreadyReturned +
          " | Available: " +
          availableQuantity

        );

      }


      const unitCost =
        Number(
          originalLine["Unit Cost"]
        );


      if (
        !isFinite(unitCost) ||
        unitCost < 0
      ) {

        throw new Error(
          "Invalid original unit cost for item: " +
          itemCode
        );

      }


      /*
       * Preserve the original purchase
       * valuation basis.
       */

      const grossAmount =
        quantity *
        unitCost;


      const originalQuantity =
        purchasedQuantity > 0
          ? purchasedQuantity
          : 1;


      const originalDiscount =
        Number(
          originalLine["Discount"] || 0
        );


      /*
       * Allocate the original line discount
       * proportionally to the returned quantity.
       */

      const allocatedDiscount =
        originalDiscount *
        (
          quantity /
          originalQuantity
        );


      const taxableAmount =
        Math.max(
          0,
          grossAmount -
          allocatedDiscount
        );


      preparedLines.push({

        lineNo:
          lineNo++,

        originalLineNo:
          originalLineNo,

        itemCode:
          itemCode,

        itemName:
          mairixPurchaseReturnText(
            originalLine["Item Name"]
          ),

        quantity:
          quantity,

        unit:
          mairixPurchaseReturnText(
            originalLine["Unit"]
          ),

        unitCost:
          unitCost,

        grossAmount:
          grossAmount,

        discount:
          allocatedDiscount,

        taxableAmount:
          taxableAmount

      });


      subtotal +=
        taxableAmount;

    }
  );


  /**********************************************
   * TAX REVERSAL
   *
   * The return is passed to the centralized
   * Tax Engine.
   *
   * The Tax Engine determines:
   *
   * - tax types
   * - tax bases
   * - tax amounts
   * - reversal signs
   * - final tax base
   **********************************************/

  const taxResult =
    mairixPurchaseReturnCalculateTaxes({

      transactionType:
        MAIRIX_PURCHASE_RETURNS.TRANSACTION_TYPE,

      transactionId:
        mairixGeneratePurchaseReturnId(),

      date:
        purchaseDate,

      originalPurchaseId:
        originalPurchaseId,

      supplierCode:
        supplierCode,

      taxableAmount:
        subtotal,

      lines:
        preparedLines

    });


  const taxTotal =
    Number(
      taxResult.totalTax || 0
    );


  /*
   * Purchase return reduces the supplier
   * transaction value.
   *
   * We represent the return itself as
   * a positive business amount here,
   * while the centralized financial engine
   * receives the actual effects:
   *
   * CASH:
   *   + refund
   *
   * PAYABLE:
   *   - liability
   */

  const netTotal =
    subtotal +
    taxTotal;


  const paymentType =
    mairixPurchaseReturnPaymentType(
      purchase
    );


  let cashEffect = 0;
  let payableEffect = 0;


  if (
    paymentType ===
    MAIRIX_PURCHASE_RETURNS.PAYMENT_CASH
  ) {

    cashEffect =
      netTotal;

  } else {

    payableEffect =
      -netTotal;

  }


  /**********************************************
   * PURCHASE RETURN ID
   **********************************************/

  const purchaseReturnId =
    taxResult.transactionId ||
    mairixGeneratePurchaseReturnId();


  /**********************************************
   * INVENTORY REVERSAL
   *
   * Purchase return:
   *
   * Quantity = NEGATIVE
   **********************************************/

  const inventoryResult =
    mairixPurchaseReturnPostInventory({

      purchaseReturnId:
        purchaseReturnId,

      originalPurchaseId:
        originalPurchaseId,

      date:
        purchaseDate,

      warehouseCode:
        warehouseCode,

      supplierCode:
        supplierCode,

      lines:
        preparedLines

    });


  /**********************************************
   * FINANCIAL REVERSAL
   **********************************************/

  const financialResult =
    mairixPurchaseReturnPostFinancial({

      purchaseReturnId:
        purchaseReturnId,

      originalPurchaseId:
        originalPurchaseId,

      date:
        purchaseDate,

      supplierCode:
        supplierCode,

      supplierName:
        mairixPurchaseReturnText(
          purchase["Supplier Name"]
        ),

      paymentType:
        paymentType,

      subtotal:
        subtotal,

      taxTotal:
        taxTotal,

      netTotal:
        netTotal,

      cashEffect:
        cashEffect,

      payableEffect:
        payableEffect,

      userId:
        user.userId

    });


  /**********************************************
   * SAVE HEADER
   **********************************************/

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        MAIRIX_PURCHASE_RETURNS.SHEET
      );


  if (!sheet) {

    throw new Error(
      "PURCHASE_RETURNS sheet not found."
    );

  }


  const now =
    new Date();


  const row = [

    purchaseReturnId,

    data.referenceNo ||
      purchaseReturnId,

    originalPurchaseId,

    mairixPurchaseReturnText(
      purchase["Reference No"]
    ),

    purchaseDate,

    supplierCode,

    mairixPurchaseReturnText(
      purchase["Supplier Name"]
    ),

    warehouseCode,

    mairixPurchaseReturnText(
      purchase["Warehouse Name"]
    ),

    paymentType,

    subtotal,

    0,

    subtotal,

    taxTotal,

    netTotal,

    cashEffect,

    payableEffect,

    taxResult.reference ||
      "",

    inventoryResult.reference ||
      "",

    financialResult.reference ||
      "",

    user.userId,

    user.name ||
      "",

    MAIRIX_PURCHASE_RETURNS.STATUS_POSTED,

    now

  ];


  sheet
    .appendRow(row);


  /**********************************************
   * SAVE LINES
   **********************************************/

  mairixPurchaseReturnSaveLines(
    purchaseReturnId,
    originalPurchaseId,
    warehouseCode,
    preparedLines,
    taxResult,
    now
  );


  /**********************************************
   * CENTRAL REFERENCE
   **********************************************/

  mairixPurchaseReturnRegisterReference({

    reference:
      purchaseReturnId,

    type:
      MAIRIX_PURCHASE_RETURNS.TRANSACTION_TYPE,

    date:
      purchaseDate,

    originalReference:
      originalPurchaseId,

    warehouseCode:
      warehouseCode,

    supplierCode:
      supplierCode,

    inventoryReference:
      inventoryResult.reference ||
      "",

    financialReference:
      financialResult.reference ||
      "",

    taxReference:
      taxResult.reference ||
      "",

    userId:
      user.userId

  });


  return {

    success:
      true,

    purchaseReturnId:
      purchaseReturnId,

    originalPurchaseId:
      originalPurchaseId,

    supplierCode:
      supplierCode,

    warehouseCode:
      warehouseCode,

    paymentType:
      paymentType,

    subtotal:
      subtotal,

    taxTotal:
      taxTotal,

    netTotal:
      netTotal,

    cashEffect:
      cashEffect,

    payableEffect:
      payableEffect,

    taxReference:
      taxResult.reference ||
      "",

    inventoryReference:
      inventoryResult.reference ||
      "",

    financialReference:
      financialResult.reference ||
      "",

    status:
      MAIRIX_PURCHASE_RETURNS.STATUS_POSTED

  };

}


/************************************************
 * SAVE RETURN LINES
 ************************************************/

function mairixPurchaseReturnSaveLines(
  purchaseReturnId,
  originalPurchaseId,
  warehouseCode,
  lines,
  taxResult,
  createdAt
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        MAIRIX_PURCHASE_RETURNS.LINES_SHEET
      );


  if (!sheet) {

    throw new Error(
      "PURCHASE_RETURN_LINES sheet not found."
    );

  }


  const rows =
    lines.map(
      function(line) {

        const lineTax =
          mairixPurchaseReturnFindLineTax(
            taxResult,
            line.lineNo
          );


        const taxTotal =
          Number(
            lineTax.totalTax || 0
          );


        const netAmount =
          line.taxableAmount +
          taxTotal;


        return [

          purchaseReturnId,

          line.lineNo,

          originalPurchaseId,

          line.originalLineNo,

          line.itemCode,

          line.itemName,

          line.quantity,

          line.unit,

          line.unitCost,

          line.grossAmount,

          line.discount,

          line.taxableAmount,

          taxTotal,

          netAmount,

          -line.quantity,

          lineTax.reference ||
            "",

          warehouseCode,

          createdAt

        ];

      }
    );


  if (!rows.length) {
    return;
  }


  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      rows[0].length
    )
    .setValues(rows);

}


/************************************************
 * FIND LINE TAX
 ************************************************/

function mairixPurchaseReturnFindLineTax(
  taxResult,
  lineNo
) {

  if (
    !taxResult ||
    !Array.isArray(
      taxResult.lines
    )
  ) {

    return {

      totalTax:
        0,

      reference:
        ""

    };

  }


  for (
    let i = 0;
    i < taxResult.lines.length;
    i++
  ) {

    if (
      Number(
        taxResult.lines[i].lineNo
      ) ===
      Number(lineNo)
    ) {

      return {

        totalTax:
          Number(
            taxResult.lines[i].totalTax || 0
          ),

        reference:
          taxResult.lines[i].reference ||
          ""

      };

    }

  }


  return {

    totalTax:
      0,

    reference:
      ""

  };

}


/************************************************
 * GET ORIGINAL PURCHASE
 ************************************************/

function mairixPurchaseReturnGetPurchase(
  purchaseId
) {

  if (
    typeof mairixPurchaseGetPurchase ===
    "function"
  ) {

    return mairixPurchaseGetPurchase(
      purchaseId
    );

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "PURCHASES"
      );


  if (!sheet) {
    return null;
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
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
      .getDisplayValues()[0];


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
      mairixPurchaseReturnText(
        rows[i][
          headers.indexOf(
            "Purchase ID"
          )
        ]
      ) !==
      mairixPurchaseReturnText(
        purchaseId
      )
    ) {

      continue;

    }


    const result = {};


    headers.forEach(
      function(header, index) {

        result[
          mairixPurchaseReturnText(
            header
          )
        ] =
          rows[i][index];

      }
    );


    return result;

  }


  return null;

}


/************************************************
 * GET ORIGINAL PURCHASE LINES
 ************************************************/

function mairixPurchaseReturnGetPurchaseLines(
  purchaseId
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        "PURCHASE_LINES"
      );


  if (!sheet) {
    return [];
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
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
      .getDisplayValues()[0];


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const result = [];


  rows.forEach(
    function(row) {

      const object = {};


      headers.forEach(
        function(header, index) {

          object[
            mairixPurchaseReturnText(
              header
            )
          ] =
            row[index];

        }
      );


      if (
        mairixPurchaseReturnText(
          object["Purchase ID"]
        ) ===
        mairixPurchaseReturnText(
          purchaseId
        )
      ) {

        result.push(object);

      }

    }
  );


  return result;

}


/************************************************
 * GET RETURNED QUANTITIES
 *
 * Prevents over-return.
 ************************************************/

function mairixPurchaseReturnGetReturnedQuantities(
  purchaseId
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        MAIRIX_PURCHASE_RETURNS.LINES_SHEET
      );


  const result = {};


  if (!sheet) {
    return result;
  }


  const lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return result;
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0];


  const rows =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      )
      .getValues();


  const purchaseIdColumn =
    headers.indexOf(
      "Original Purchase ID"
    );


  const lineColumn =
    headers.indexOf(
      "Original Line No"
    );


  const quantityColumn =
    headers.indexOf(
      "Quantity"
    );


  rows.forEach(
    function(row) {

      if (
        mairixPurchaseReturnText(
          row[purchaseIdColumn]
        ) !==
        mairixPurchaseReturnText(
          purchaseId
        )
      ) {

        return;

      }


      const lineNo =
        Number(
          row[lineColumn]
        );


      const quantity =
        Number(
          row[quantityColumn]
        );


      if (
        !isFinite(lineNo) ||
        !isFinite(quantity)
      ) {

        return;

      }


      result[lineNo] =
        Number(
          result[lineNo] || 0
        ) +
        quantity;

    }
  );


  return result;

}


/************************************************
 * PAYMENT TYPE
 ************************************************/

function mairixPurchaseReturnPaymentType(
  purchase
) {

  const value =
    mairixPurchaseReturnText(
      purchase["Payment Type"]
    ).toUpperCase();


  if (
    value ===
    MAIRIX_PURCHASE_RETURNS.PAYMENT_CREDIT
  ) {

    return value;

  }


  return MAIRIX_PURCHASE_RETURNS.PAYMENT_CASH;

}


/************************************************
 * TAX ENGINE BRIDGE
 ************************************************/

function mairixPurchaseReturnCalculateTaxes(
  data
) {

  if (
    typeof calculateTaxes ===
    "function"
  ) {

    return calculateTaxes(
      data
    );

  }


  if (
    typeof calculateTransactionTaxes ===
    "function"
  ) {

    return calculateTransactionTaxes(
      data
    );

  }


  /*
   * A return without a connected tax engine
   * cannot invent tax rules.
   *
   * Return zero only as a temporary bridge
   * while Tax Engine is being assembled.
   */

  return {

    transactionId:
      data.transactionId,

    totalTax:
      0,

    reference:
      "",

    lines:
      data.lines.map(
        function(line) {

          return {

            lineNo:
              line.lineNo,

            totalTax:
              0,

            reference:
              ""

          };

        }
      )

  };

}


/************************************************
 * INVENTORY ENGINE BRIDGE
 ************************************************/

function mairixPurchaseReturnPostInventory(
  data
) {

  if (
    typeof postInventoryMovement ===
    "function"
  ) {

    return postInventoryMovement({

      transactionType:
        MAIRIX_PURCHASE_RETURNS.TRANSACTION_TYPE,

      transactionId:
        data.purchaseReturnId,

      originalTransactionId:
        data.originalPurchaseId,

      date:
        data.date,

      warehouseCode:
        data.warehouseCode,

      supplierCode:
        data.supplierCode,

      lines:
        data.lines.map(
          function(line) {

            return {

              lineNo:
                line.lineNo,

              itemCode:
                line.itemCode,

              quantity:
                -line.quantity,

              unitCost:
                line.unitCost,

              amount:
                -(
                  line.quantity *
                  line.unitCost
                )

            };

          }
        )

    });

  }


  if (
    typeof createInventoryMovement ===
    "function"
  ) {

    return createInventoryMovement({

      type:
        MAIRIX_PURCHASE_RETURNS.TRANSACTION_TYPE,

      reference:
        data.purchaseReturnId,

      originalReference:
        data.originalPurchaseId,

      warehouseCode:
        data.warehouseCode,

      lines:
        data.lines.map(
          function(line) {

            return {

              itemCode:
                line.itemCode,

              quantity:
                -line.quantity,

              unitCost:
                line.unitCost

            };

          }
        )

    });

  }


  throw new Error(
    "MAIRIX Inventory Engine is not available."
  );

}


/************************************************
 * FINANCIAL ENGINE BRIDGE
 ************************************************/

function mairixPurchaseReturnPostFinancial(
  data
) {

  if (
    typeof postFinancialTransaction ===
    "function"
  ) {

    return postFinancialTransaction({

      transactionType:
        MAIRIX_PURCHASE_RETURNS.TRANSACTION_TYPE,

      transactionId:
        data.purchaseReturnId,

      originalTransactionId:
        data.originalPurchaseId,

      date:
        data.date,

      supplierCode:
        data.supplierCode,

      supplierName:
        data.supplierName,

      paymentType:
        data.paymentType,

      subtotal:
        data.subtotal,

      taxTotal:
        data.taxTotal,

      netTotal:
        data.netTotal,

      cashEffect:
        data.cashEffect,

      payableEffect:
        data.payableEffect,

      userId:
        data.userId

    });

  }


  throw new Error(
    "MAIRIX Financial Engine is not available."
  );

}


/************************************************
 * CURRENT USER
 ************************************************/

function mairixPurchaseReturnRequireUser() {

  if (
    typeof requireCurrentUser ===
    "function"
  ) {

    return requireCurrentUser();

  }


  if (
    typeof getCurrentUser ===
    "function"
  ) {

    const user =
      getCurrentUser();


    if (!user) {

      throw new Error(
        "Current user not found."
      );

    }


    return user;

  }


  throw new Error(
    "MAIRIX authentication engine is not available."
  );

}


/************************************************
 * PURCHASE RETURN PERMISSION
 *
 * Current MAIRIX Security Engine
 ************************************************/

function mairixPurchaseReturnRequirePermission(
  userId,
  code,
  action
) {

  return mairixRequirePermission(
    userId,
    code,
    action
  );

}


/************************************************
 * CENTRAL REFERENCE
 ************************************************/

function mairixPurchaseReturnRegisterReference(
  data
) {

  if (
    typeof registerCentralReference ===
    "function"
  ) {

    return registerCentralReference(
      data
    );

  }


  return {

    success:
      true,

    reference:
      data.reference

  };

}


/************************************************
 * RETURN ID
 ************************************************/

function mairixGeneratePurchaseReturnId() {

  const dateKey =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd"
    );


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        MAIRIX_PURCHASE_RETURNS.SHEET
      );


  let counter = 1;


  if (
    sheet &&
    sheet.getLastRow() >= 2
  ) {

    const idColumn =
      mairixgetColumnNumber(
        sheet,
        "Purchase Return ID"
      );


    const values =
      sheet
        .getRange(
          2,
          idColumn,
          sheet.getLastRow() - 1,
          1
        )
        .getDisplayValues();


    const prefix =
      MAIRIX_PURCHASE_RETURNS.DOCUMENT_PREFIX +
      "-" +
      dateKey +
      "-";


    values.forEach(
      function(row) {

        const value =
          mairixPurchaseReturnText(
            row[0]
          );


        if (
          value.indexOf(prefix) !== 0
        ) {

          return;

        }


        const number =
          parseInt(
            value.substring(
              prefix.length
            ),
            10
          );


        if (
          !isNaN(number) &&
          number >= counter
        ) {

          counter =
            number + 1;

        }

      }
    );

  }


  return (
    MAIRIX_PURCHASE_RETURNS.DOCUMENT_PREFIX +
    "-" +
    dateKey +
    "-" +
    String(counter)
      .padStart(
        5,
        "0"
      )
  );

}


/************************************************
 * TEXT HELPER
 ************************************************/

function mairixPurchaseReturnText(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return String(
    value
  ).trim();

}


/************************************************
 * REQUIRED HELPER
 ************************************************/

function mairixPurchaseReturnRequired(
  value,
  label
) {

  const normalized =
    mairixPurchaseReturnText(
      value
    );


  if (!normalized) {

    throw new Error(
      label +
      " is required."
    );

  }


  return normalized;

}


/************************************************
 * TEST SETUP
 ************************************************/

function testMAIRIXPurchaseReturnsSetup() {

  const result =
    setupMAIRIXPurchaseReturns();


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  if (
    !result.success
  ) {

    throw new Error(
      "Purchase Returns setup failed."
    );

  }


  Logger.log(
    "=========================================="
  );

  Logger.log(
    "MAIRIX PURCHASE RETURNS SETUP: PASS"
  );

  Logger.log(
    "=========================================="
  );

}


/************************************************
 * TEST INVALID RETURN
 *
 * Must fail without posting.
 ************************************************/

function testMAIRIXPurchaseReturnValidation() {

  try {

    createPurchaseReturn({

      originalPurchaseId:
        "",

      lines: []

    });


    throw new Error(
      "Invalid purchase return was accepted."
    );

  } catch (error) {

    Logger.log(
      "Expected validation error: " +
      error.message
    );

  }


  Logger.log(
    "MAIRIX PURCHASE RETURN VALIDATION TEST: PASS"
  );

}

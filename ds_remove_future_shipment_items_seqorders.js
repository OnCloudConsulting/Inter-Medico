/**
 * @NApiVersion 2.1
 * @NScriptType UserEventScript
 */
define(['N/search'], (search) => {

  const beforeSubmit = (ctx) => {
    if (ctx.type !== ctx.UserEventType.CREATE) return;

    const rec = ctx.newRecord;
    const soId = rec.getValue('createdfrom');
    if (!soId) return;

    // Eligible SO lines: release date today or earlier AND committed qty > 0
    const eligible = {}; // lineId -> committedQty
    search.create({
      type: search.Type.SALES_ORDER,
      filters: [
        ['internalid', 'anyof', soId], 'AND',
        ['mainline', 'is', 'F'], 'AND',
        ['taxline', 'is', 'F'], 'AND',
        ['shipping', 'is', 'F'], 'AND',
        ['cogs', 'is', 'F'], 'AND',
        ['custcol_release_date_for_picking', 'onorbefore', 'today'], 'AND',
        ['formulanumeric: NVL({quantitycommitted},0)', 'greaterthan', 0]
      ],
      columns: ['line', 'quantitycommitted']
    }).run().each(r => {
      eligible[String(r.getValue('line'))] = Number(r.getValue('quantitycommitted')) || 0;
      return true;
    });

    let kept = 0;
    const count = rec.getLineCount({ sublistId: 'item' });

    for (let i = 0; i < count; i++) {
      if (!rec.getSublistValue({ sublistId: 'item', fieldId: 'itemreceive', line: i })) continue;

      const orderLine = String(rec.getSublistValue({ sublistId: 'item', fieldId: 'orderline', line: i }));
      const committedQty = eligible[orderLine];

      if (committedQty > 0) {
        // Don't fulfill more than what's committed
        const qty = Number(rec.getSublistValue({ sublistId: 'item', fieldId: 'quantity', line: i })) || 0;
        if (qty > committedQty) {
          rec.setSublistValue({ sublistId: 'item', fieldId: 'quantity', line: i, value: committedQty });
        }
        kept++;
      } else {
        // Future release date, blank date, or nothing committed
        rec.setSublistValue({ sublistId: 'item', fieldId: 'itemreceive', line: i, value: false });
      }
    }

    if (kept === 0) {
      throw 'No lines on this order are released for picking today with committed quantity.';
    }
  };

  return { beforeSubmit };
});

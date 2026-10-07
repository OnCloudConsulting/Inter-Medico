/**
 * @NApiVersion 2.1
 * @NScriptType ClientScript
 */
define(['N/search', 'N/log'], (search, log) => {

    const TAG = '[DUP_PO_CHECK]';

    // Logs to browser console (F12) AND script execution log
    const dbg = (title, details) => {
        try {
            console.log(TAG + ' ' + title, details !== undefined ? details : '');
        } catch (ignore) {}
        try {
            log.debug({ title: TAG + ' ' + title, details: details });
        } catch (ignore) {}
    };

    // Fires on form load - confirms script is attached/deployed at all
    const pageInit = (context) => {
        dbg('pageInit fired', {
            mode: context.mode,
            recordType: context.currentRecord.type,
            recordId: context.currentRecord.id
        });
    };

    const saveRecord = (context) => {
        dbg('saveRecord fired');

        try {
            const rec = context.currentRecord;

            dbg('Record info', { type: rec.type, id: rec.id });

            const rawPo = rec.getValue({ fieldId: 'otherrefnum' });
            dbg('Raw otherrefnum value', { value: rawPo, jsType: typeof rawPo });

            const po = String(rawPo || '').trim();

            if (!po) {
                dbg('PO is empty - skipping check, allowing save');
                return true;
            }

            const filters = [
                ['mainline', 'is', 'T'],
                'AND',
                ['otherrefnum', 'equalto', po]
            ];

            if (rec.id) {
                filters.push('AND', ['internalid', 'noneof', rec.id]);
            }

            dbg('Search filters', JSON.stringify(filters));

            const salesOrderSearch = search.create({
                type: search.Type.SALES_ORDER,
                filters: filters,
                columns: [
                    search.createColumn({ name: 'internalid' }),
                    search.createColumn({ name: 'tranid' }),
                    search.createColumn({ name: 'otherrefnum' })
                ]
            });

            const results = salesOrderSearch.run().getRange({ start: 0, end: 1 });

            dbg('Search result count', results.length);

            if (results.length > 0) {
                const soNumber = results[0].getValue({ name: 'tranid' });
                const soId = results[0].getValue({ name: 'internalid' });

                dbg('Duplicate found', { soId: soId, soNumber: soNumber, po: po });

                alert(
                    'Duplicate PO Number Found.\n\n' +
                    'PO # ' + po +
                    ' already exists on Sales Order ' + soNumber +
                    '.\n\nPlease enter a different PO number.'
                );

                return false;
            }

            dbg('No duplicate found - allowing save');
            return true;

        } catch (e) {
            // Previously swallowed silently - this is likely why it looked like nothing ran
            dbg('ERROR in saveRecord', {
                name: e.name,
                message: e.message,
                stack: e.stack
            });
            console.error(TAG + ' ERROR', e);

            // Allow save if unexpected script error occurs
            return true;
        }
    };

    return {
        pageInit: pageInit,
        saveRecord: saveRecord
    };
});

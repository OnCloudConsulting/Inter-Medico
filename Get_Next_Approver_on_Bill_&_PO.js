/**
 * @NApiVersion 2.x
 * @NScriptType workflowactionscript
 */
define(['N/search'], function (search) {

  function onAction(scriptContext) {
    var rec = scriptContext.newRecord;

    // Default: look up the expense report's employee
    var lookupEmp = rec.getValue('entity');

    // If a next approver already exists, look up THAT person instead
    var nextApprover = rec.getValue('nextapprover');
    log.debug({ title: 'entity / nextapprover', details: lookupEmp + ' / ' + nextApprover });

    if (!isEmpty(nextApprover) && String(nextApprover) !== '-1') {
      lookupEmp = nextApprover;
    }

    if (isEmpty(lookupEmp)) {
      log.error({ title: 'No employee to look up', details: 'Record ID: ' + rec.id });
      return '';
    }

    var emp = search.lookupFields({
      type: search.Type.EMPLOYEE,
      id: lookupEmp,
      columns: ['approver', 'supervisor'] // approver = Expense Approver
    });
    log.debug({ title: 'lookup for ' + lookupEmp, details: emp });

    // 1st priority: Expense Approver
    if (emp.approver && emp.approver.length > 0) {
      log.debug({ title: 'Next approver (Expense Approver)', details: emp.approver[0].value });
      return emp.approver[0].value;
    }

    // 2nd priority: Supervisor
    if (emp.supervisor && emp.supervisor.length > 0) {
      log.debug({ title: 'Next approver (Supervisor)', details: emp.supervisor[0].value });
      return emp.supervisor[0].value;
    }

    log.error({ title: 'No Expense Approver or Supervisor', details: 'Employee ' + lookupEmp + ', Record ID: ' + rec.id });
    return '';
  }

  function isEmpty(value) {
    return value === null || value === undefined || value === '' ||
           value === ' ' || value === 'null';
  }

  return { onAction: onAction };
});
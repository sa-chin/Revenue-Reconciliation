// this script automates the revenue reconciliation calculations.
// start with a raw download from here: https://platform.datorama.com/108713/visualize/11267125/page/v2/4647472
// import (or copy/paste) this into google sheets
// delete rows with 'not valid', 'unknown', etc
// do not change any column names


// creates menu button
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Sachin Things')
    .addItem('Find Contracted Actual', 'mapOD')
    .addToUi();
}


// starts function, copies data to use without affecting the original 
function mapOD() {
  const ss = SpreadsheetApp.getActiveSheet();
  const data = ss.getDataRange().getValues();
  const header = data[0].slice();
  header.push('Weight', 'Cost per KPI', 'Contracted Actual', 'Extra Delivery', 'Extra Spend');
  const records = data.slice(1);

  // groups rows by Op1 Line ID, sums total delivery
  const groupedLineIDs = {};
  records.forEach(record => {
      const lineID = record[header.indexOf('*Operative Line Item ID')];

      if (!groupedLineIDs[lineID]) {
        groupedLineIDs[lineID] = {
          totalDelivery: 0,
          contractedGoal: record[header.indexOf('Contracted Quantity (Operative)')],
          rows: []
        };
      };

      groupedLineIDs[lineID].rows.push(record);
      groupedLineIDs[lineID].totalDelivery += Number(record[header.indexOf('*KPI Value')]);
    }
  );

  // process the grouped data
  // for each grouping, for each row, find the % distribution, weighted goal, extra delivery, and extra spend
  Object.values(groupedLineIDs).forEach(group => {
    const { totalDelivery, contractedGoal, rows } = group;

    rows.forEach(row => {
      const kpi = Number(row[header.indexOf('*KPI Value')]) || 0;
      const spend = Number(row[header.indexOf('Amount Spent / Media Cost')]) || 0;
      let costPerKPI = 0

      switch (String(row[header.indexOf('Cost Type (Operative)')])) {
        case 'CPM':
          costPerKPI = kpi ? (spend / kpi) * 1000 : 0;
          break;
        case 'CPC':
          costPerKPI = kpi ? (spend / kpi) : 0;
          break;
        case 'Cost Per Unit':
          costPerKPI = kpi ? (spend / kpi) : 0;
          break;
        default:
          break;
      }

      const weight = kpi / totalDelivery;
      const contractedActual = weight * contractedGoal;
      const extraDelivery = kpi - contractedActual;
      let extraSpend = 0;

      switch (String(row[header.indexOf('Cost Type (Operative)')])) {
        case 'CPM':
          extraSpend = extraDelivery / 1000 * costPerKPI;
          break;
        case 'CPC':
          extraSpend = extraDelivery * costPerKPI;
          break;
        case 'Cost Per Unit':
          extraSpend = extraDelivery * costPerKPI;
          break;
        default:
          break;
      }

      row.push(weight, costPerKPI, contractedActual, extraDelivery, extraSpend);
    });
  });

  // put the processed data into the sheet: to be completed
  const out = [header, ...records];
  ss.clearContents();
  ss.getRange(1, 1, out.length, header.length).setValues(out);
}
// ==================================================
// Revenue Reconciliation Automation
// ==================================================

// ---- Ingested headers (do not rename in sheet) ----
const Header_Region = '*Region';
const Header_Advertiser = '*Advertiser';
const Header_Platform = '*Platform';
const Header_CampaignName = 'Campaign Name';
const Header_Op1SalesID = '*Operative ID';
const Header_AdSetName = 'Ad Set Name | Line Item Name (GAM)';
const Header_Op1LineID = '*Operative Line Item ID';
const Header_CostType = 'Cost Type (Operative)';
const Header_ContractedQuantity = 'Contracted Quantity (Operative)';
const Header_Impressions = '*Impressions';
const Header_VideoViews = 'Video Views';
const Header_Clicks = '*Clicks (All)';
const Header_KPIValue = '*KPI Value';
const Header_Spend = 'Amount Spent / Media Cost';

// ---- Calculated columns ----
const Col_Weight = "Weight";
const Col_CostPerKPI = 'Cost per KPI';
const Col_ContractedActual = 'Contracted Actual';
const Col_ExtraDelivery = 'Extra Delivery';
const Col_ExtraSpend = 'Extra Spend';

// ==================================================
// Menu button
// ==================================================
function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Sachin Things')
    .addItem('Find Contracted Actual', 'mapOverDelivery')
    .addToUi();
}

// ==================================================
// Helpers
// ==================================================
function calculateCostPerKPI(costType, spend, kpi) {
  if (!kpi) return 0;
  switch (costType) {
    case 'CPM': return (spend / kpi) * 1000;
    case 'CPC':
    case 'Cost Per Unit': return spend / kpi;
    default: return 0;
  }
}

function calculateExtraSpend(costType, extraDelivery, costPerKPI) {
  switch (costType) {
    case 'CPM': return (extraDelivery / 1000) * costPerKPI;
    case 'CPC':
    case 'Cost Per Unit': return extraDelivery * costPerKPI;
    default: return 0;
  }
}

function validateHeaders(headerRow) {
  const requiredHeaders = [
    Header_Region, Header_Advertiser, Header_Platform,
    Header_CampaignName, Header_Op1SalesID, Header_AdSetName,
    Header_Op1LineID, Header_CostType, Header_ContractedQuantity,
    Header_Impressions, Header_VideoViews, Header_Clicks,
    Header_KPIValue, Header_Spend,
  ];

  requiredHeaders.forEach(h => {
    if (headerRow.indexOf(h) === -1) {
      throw new Error(`Missing column: ${h}`);
    }
  });
}

// ==================================================
// Main function
// ==================================================
function mapOverDelivery() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No data found.");
    return;
  }

  const header = data[0].slice();
  validateHeaders(header);

  // Build column index map
  const colIdx = {
    region: header.indexOf(Header_Region),
    advertiser: header.indexOf(Header_Advertiser),
    platform: header.indexOf(Header_Platform),
    campaignName: header.indexOf(Header_CampaignName),
    op1SalesID: header.indexOf(Header_Op1SalesID),
    adSetName: header.indexOf(Header_AdSetName),
    op1LineID: header.indexOf(Header_Op1LineID),
    costType: header.indexOf(Header_CostType),
    contractedQuantity: header.indexOf(Header_ContractedQuantity),
    impressions: header.indexOf(Header_Impressions),
    videoViews: header.indexOf(Header_VideoViews),
    clicks: header.indexOf(Header_Clicks),
    kpiValue: header.indexOf(Header_KPIValue),
    spend: header.indexOf(Header_Spend),
  };

  // Extend header with calculated columns
  header.push(
    Col_Weight,
    Col_CostPerKPI,
    Col_ContractedActual,
    Col_ExtraDelivery,
    Col_ExtraSpend
  );

  const records = data.slice(1);

  // ---- Group by Line ID ----
  const groupedLineIDs = {};
  records.forEach(row => {
    const lineID = row[colIdx.op1LineID];

    if (!groupedLineIDs[lineID]) {
      groupedLineIDs[lineID] = {
        totalDelivery: 0,
        contractedGoal: row[colIdx.contractedQuantity],
        rows: []
      };
    }

    groupedLineIDs[lineID].rows.push(row);
    groupedLineIDs[lineID].totalDelivery += Number(row[colIdx.kpiValue]) || 0;
  });

  // ---- Process each group ----
  Object.values(groupedLineIDs).forEach(({ totalDelivery, contractedGoal, rows }) => {
    rows.forEach(row => {
      const kpi = Number(row[colIdx.kpiValue]) || 0;
      const spend = Number(row[colIdx.spend]) || 0;
      const costType = String(row[colIdx.costType]);

      const costPerKPI = calculateCostPerKPI(costType, spend, kpi);
      const weight = totalDelivery ? kpi / totalDelivery : 0;
      const contractedActual = weight * contractedGoal;
      const extraDelivery = kpi - contractedActual;
      const extraSpend = calculateExtraSpend(costType, extraDelivery, costPerKPI);

      row.push(weight, costPerKPI, contractedActual, extraDelivery, extraSpend);
    });
  });

  // ---- Write results ----
  const out = [header, ...records];

  // Clear content only (preserves formatting)
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();

  // Write new results
  sheet.getRange(1, 1, out.length, header.length).setValues(out);
}

import XLSX from 'xlsx';

export function generateTenantWorkbook(tenant, rawData) {
  const { candidates, agents, agencies, medicals, mofas, visas, flights } = rawData;

  const agentMap = new Map(agents?.map(a => [a.id, a.name]) || []);
  const agencyMap = new Map(agencies?.map(ag => [ag.id, ag.name]) || []);
  const candidateMap = new Map(candidates?.map(c => [c.id, c]) || []);

  const formattedCandidates = (candidates || []).map(c => ({
    "SL": c.sl,
    "Candidate Name": c.name,
    "Passport No": c.passport_no,
    "Current Stage": c.current_stage,
    "Agent Name": c.agent_id ? (agentMap.get(c.agent_id) || "Unknown Agent") : "N/A",
    "Received Date": c.received_date,
    "Is Returned": c.is_returned ? "Yes" : "No",
    "Returned Date": c.returned_date || "N/A"
  }));

  const formattedAgents = (agents || []).map(a => ({
    "SL": a.sl,
    "Agent Code": a.code,
    "Agent Name": a.name,
    "Status": a.is_active ? "Active" : "Inactive"
  }));

  const formattedAgencies = (agencies || []).map(ag => ({
    "SL": ag.sl,
    "Agency Code": ag.code,
    "Agency Name": ag.name,
    "Phone": ag.phone,
    "Email": ag.email
  }));

  const formattedMedicals = (medicals || []).map(m => {
    const cand = candidateMap.get(m.candidate_id);
    return {
      "Candidate SL": cand?.sl || "N/A",
      "Candidate Name": cand?.name || "Unknown",
      "Passport No": cand?.passport_no || "N/A",
      "Medical Date": m.medical_date,
      "Fit Date": m.fit_date,
      "Status": m.status
    };
  });

  const formattedMofas = (mofas || []).map(mf => {
    const cand = candidateMap.get(mf.candidate_id);
    return {
      "SL": mf.sl,
      "Application No": mf.application_number,
      "Candidate Name": cand?.name || "Unknown",
      "Passport No": cand?.passport_no || "N/A",
      "Trade": mf.trade,
      "Agency Name": mf.agency_id ? (agencyMap.get(mf.agency_id) || "N/A") : "N/A",
      "Stage": mf.stage,
      "Application Date": mf.application_date
    };
  });

  const formattedVisas = (visas || []).map(v => {
    const cand = candidateMap.get(v.candidate_id);
    return {
      "SL": v.sl,
      "Visa No": v.visa_no,
      "Candidate Name": cand?.name || "Unknown",
      "Passport No": cand?.passport_no || "N/A",
      "Visa Type": v.visa_type,
      "Status": v.status,
      "Agency Name": v.agency_id ? (agencyMap.get(v.agency_id) || "N/A") : "N/A",
      "Visa Date": v.visa_date,
      "Expiry Date": v.expiry_date
    };
  });

  const formattedFlights = (flights || []).map(f => {
    const cand = candidateMap.get(f.candidate_id);
    return {
      "SL": f.sl,
      "Candidate Name": cand?.name || "Unknown",
      "Passport No": cand?.passport_no || "N/A",
      "Flight No": f.flight_no,
      "Airline": f.airline,
      "Departure City": f.departure_city,
      "Arrival City": f.arrival_city,
      "Flight Date": f.flight_date,
      "Status": f.status
    };
  });

  const workbook = XLSX.utils.book_new();

  const dashboard = [
    { Metric: "Tenant Name", Value: tenant.name },
    { Metric: "Slug", Value: tenant.slug },
    { Metric: "Total Candidates", Value: candidates?.length || 0 },
    { Metric: "Total Agents", Value: agents?.length || 0 },
    { Metric: "Total Agencies", Value: agencies?.length || 0 },
    { Metric: "Total Medicals", Value: medicals?.length || 0 },
    { Metric: "Total Mofas", Value: mofas?.length || 0 },
    { Metric: "Total Visas", Value: visas?.length || 0 },
    { Metric: "Total Flights", Value: flights?.length || 0 },
    { Metric: "Backup Export Time", Value: new Date().toISOString() }
  ];

  // Helper function to apply styling (Font size 12, text wrapping, and auto column widths)
  function addStyledSheet(wb, dataArray, sheetName) {
    if (!dataArray || dataArray.length === 0) return;
    
    const worksheet = XLSX.utils.json_to_sheet(dataArray);
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    // Column widths auto-adjustment
    const colWidths = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let maxLen = 10;
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = worksheet[cellAddress];
        if (cell && cell.v) {
          const len = String(cell.v).length;
          if (len > maxLen) maxLen = len;
        }
      }
      colWidths.push({ wch: Math.min(maxLen + 4, 30) }); // Max width limit 30
    }
    worksheet['!cols'] = colWidths;

    // Apply font size 12, text wrapping, and styling to all cells
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!worksheet[cellAddress]) continue;

        // Initialize cell style object if not present
        if (!worksheet[cellAddress].s) worksheet[cellAddress].s = {};

        // Font Size 12 styling
        worksheet[cellAddress].s.font = {
          name: "Segoe UI",
          sz: 12,
          bold: R === 0 ? true : false // Header bold
        };

        // Text Wrap Enabled
        worksheet[cellAddress].s.alignment = {
          wrapText: true,
          vertical: "center",
          horizontal: R === 0 ? "center" : "left"
        };

        // Header Styling (Row 0)
        if (R === 0) {
          worksheet[cellAddress].s.fill = {
            fgColor: { rgb: "1F4E78" } // Professional Dark Blue Background
          };
          worksheet[cellAddress].s.font.color = { rgb: "FFFFFF" }; // White text
        }
      }
    }

    XLSX.utils.book_append_sheet(wb, worksheet, sheetName);
  }

  // Append sheets with modern styling
  addStyledSheet(workbook, dashboard, "Dashboard");
  addStyledSheet(workbook, formattedCandidates, "Candidates");
  addStyledSheet(workbook, formattedAgents, "Agents");
  addStyledSheet(workbook, formattedAgencies, "Agencies");
  addStyledSheet(workbook, formattedMedicals, "Medicals");
  addStyledSheet(workbook, formattedMofas, "Mofas");
  addStyledSheet(workbook, formattedVisas, "Visas");
  addStyledSheet(workbook, formattedFlights, "Flights");

  return workbook;
}
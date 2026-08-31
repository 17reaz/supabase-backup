import { createClient } from '@supabase/supabase-js';
import XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

const supabaseUrl = process.env.SUPABASE_URL;
const serviceKey = process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!supabaseUrl || !serviceKey) {
  console.error("❌ Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY environment variables.");
  process.exit(1);
}

const supabase = createClient(supabaseUrl, serviceKey);

async function runBackup() {
  console.log("🚀 Starting professional relational tenant-wise backup process...");

  // ১. সব টেন্যান্ট ফেচ করা
  const { data: tenants, error: tenantError } = await supabase
    .from('tenants')
    .select('id, name, slug');

  if (tenantError) {
    console.error("❌ Error fetching tenants:", tenantError.message);
    return;
  }

  // ইউনিক টাইমস্ট্যাম্প ফোল্ডার তৈরি (যেমন: backups/backup_2026-08-29T18-30-00-000Z)
  const now = new Date();
  const timestamp = now.toISOString().replace(/[:.]/g, '-');
  const backupDir = path.join(process.cwd(), 'backups', `backup_${timestamp}`);
  
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }
  console.log(`📂 Created latest backup directory: ${backupDir}`);

  // ২. প্রতিটি টেন্যান্টের জন্য লুপ চালিয়ে প্রফেশনাল ডেটা প্রসেসিং
  for (const tenant of tenants) {
    console.log(`\n📦 Processing Tenant: ${tenant.name} (${tenant.slug})`);
    const tenantId = tenant.id;

    try {
      const [
        { data: candidates },
        { data: agents },
        { data: agencies },
        { data: medicals },
        { data: mofas },
        { data: visas },
        { data: flights }
      ] = await Promise.all([
        supabase.from('candidates').select('*').eq('tenant_id', tenantId).eq('is_deleted', false),
        supabase.from('agents').select('*').eq('tenant_id', tenantId).eq('is_deleted', false),
        supabase.from('agencies').select('*').eq('tenant_id', tenantId),
        supabase.from('medicals').select('*').eq('tenant_id', tenantId),
        supabase.from('mofas').select('*').eq('tenant_id', tenantId),
        supabase.from('visas').select('*').eq('tenant_id', tenantId),
        supabase.from('flights').select('*').eq('tenant_id', tenantId),
      ]);

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
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dashboard), "Dashboard");

      if (formattedCandidates.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedCandidates), "Candidates");
      if (formattedAgents.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedAgents), "Agents");
      if (formattedAgencies.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedAgencies), "Agencies");
      if (formattedMedicals.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedMedicals), "Medicals");
      if (formattedMofas.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedMofas), "Mofas");
      if (formattedVisas.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedVisas), "Visas");
      if (formattedFlights.length) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(formattedFlights), "Flights");

      const filePath = path.join(backupDir, `${tenant.slug}_backup.xlsx`);
      XLSX.writeFile(workbook, filePath);
      console.log(`   ✅ Saved: ${filePath}`);

    } catch (err) {
      console.error(`   ❌ Failed for tenant ${tenant.name}:`, err.message);
    }
  }

  console.log("\n🎉 All professional tenant backups completed successfully!");
}

runBackup();
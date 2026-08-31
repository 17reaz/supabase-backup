import fs from 'fs';
import path from 'path';
import XLSX from 'xlsx';
import { supabase } from './src/config/supabase.js';
import { generateTenantWorkbook } from './src/utils/excelHelper.js';

async function runBackup() {
  console.log("🚀 Starting professional relational tenant-wise backup process...");

  const { data: tenants, error: tenantError } = await supabase
    .from('tenants')
    .select('id, name, slug');

  if (tenantError) {
    console.error("❌ Error fetching tenants:", tenantError.message);
    return;
  }

  // Day-wise & Time-wise Dynamic Folder Structure Creation
  const now = new Date();
  const dateStr = now.toISOString().split('T')[0]; // e.g., 2026-08-29
  const timeStr = now.toTimeString().split(' ')[0].replace(/:/g, '-'); // e.g., 18-30-00
  
  const backupDir = path.join(process.cwd(), 'backups', dateStr, timeStr);
  
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }
  console.log(`📂 Created latest backup directory: ${backupDir}`);

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

      const workbook = generateTenantWorkbook(tenant, {
        candidates, agents, agencies, medicals, mofas, visas, flights
      });

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